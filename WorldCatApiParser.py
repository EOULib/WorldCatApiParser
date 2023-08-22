
#Built-in/Generic Imports
import tkinter as tk #For GUI Class
import csv #For Csv_Parser Class
import sqlite3 #For Database Class
import threading #For GUI class to run the status bar in a seperate thread
import pandas as pd #For Csv_Paraser Class to check that the import file has an ISBN Header

#Libs
from urllib.request import Request, urlopen #For Csv_Parser Class to make the API Call
from urllib.error import URLError, HTTPError #For Csv_Parser Class to check for API Call Errors
import xml.etree.ElementTree as ET #For Csv_Parser Class for grabbing xml data returned by API Call
from tkinter import filedialog #For GUI Class to open file/directory explorer window
from tkinter import messagebox #For GUI Class and Csv_Parser Class to create Error popups
from tkinter.messagebox import askyesno #For GUI Class to exit program and save key and instituional code
from tkinter import ttk #For GUI Progess Bar
from tkinter.constants import DISABLED, NORMAL #For GUI Class to enable and disable the submit button


__author__="Jeremiah Kellogg"
__copyright__="Copyright 2023, Eastern Oregon University"
__credits__="Jeremiah Kellogg"
__license__="MIT License"
__version__="1.0.0"
__maintainer__="Jeremiah Kellogg"
__email__="jkellogg@eou.edu"
__status__="beta"

############# CLASSES #############

# Csv_Parser Class inputs csv file with ISBNs, then makes API Calls to OCLC for each ISBN to grab XML data
# and outputs it into a new csv file
class Csv_Parser:
    def __init__(self):
        #Global Variables
        self.message_ns = '{http://www.loc.gov/zing/srw/diagnostic/}' #namespace for record not found messages
        self.namespace = '{http://www.loc.gov/MARC21/slim}' #namespace for records found
        self.record_status = ''
        self.title = ''
        self.author = ''
        self.publisher = ''
        self.pub_year = ''
        self.edition = ''
        self.price = ''
        self.vendor_name = ''
        self.binding = ''
        self.oclc_number = ''
        self.held_by_institution= ''

        self.count = 0

    ######### Csv_Parser Methods #########

    #Open csv with ISBNs to read from
    def open_import_csv(self, name):
        self.file_name = name
        try:
            self.book_dictionary = csv.DictReader(open(self.file_name))
            if self.check_isbn_header():
                return True
            else:
                messagebox.showinfo("No ISBN Column", "Check the import csv file and ensure there is a column labeled ISBN")
                return False
        except FileNotFoundError as err:
            messagebox.showinfo("ERROR", "File or Directory Wasn't Found, Click OK and Add a Valid Import File Path")
            return False

    #Check to make sure imported csv has a column labled ISBN
    def check_isbn_header(self):
        df = pd.read_csv(self.file_name)
        list_of_column_names = list(df.columns)
        uppercase_list = [x.upper() for x in list_of_column_names]
        column_name = 'ISBN'
        if column_name in uppercase_list:
            return True
        else:
            return False

    #Create and open output csv file to write to
    def open_export_csv(self, directory, name):
        self.output_file_name_path = directory+"/"+name
        try:
            self.output_file = open(self.output_file_name_path, 'w', encoding='utf-8-sig')#Must explicitely set encoding to utf-8-sig so that Windows doesn't throw a hissy fit
        except FileNotFoundError as err:
            messagebox.showinfo("ERROR", "Directory Wasn't Found, Click OK and Add a Valid Export Directory Path")
            return False
        self.csvwriter = csv.writer(self.output_file)
        self.column_header = ['Record Status', 'ISBN', 'Title', 'Author', 'Publisher', 'Year Published', 'Edition', 'Price', 'Price Obtained From', 'Binding', "Held at Institution"]
        self.csvwriter.writerow(self.column_header)
        return True

    def get_oclc_data(self, oclc_key, oclc_code):
        self.book_dictionary_upper = [{key.upper(): value for key, value in r.items()} for r in self.book_dictionary]
        for row in self.book_dictionary_upper:
            isbn = str(row["ISBN"])
            api_key = oclc_key.get()
            url = "https://worldcat.org/webservices/catalog/content/isbn/"+isbn+"?servicelevel=full&wskey="+api_key
            api_call = Request(url, headers={"User-Agent": "Mozilla/5.0"})
            raw_data = urlopen(api_call).read()
            root = ET.fromstring(raw_data)
            record_found = True

            #Check to make sure a record is available through OCLC before trying to parse any data
            for element in root.iter(self.message_ns+'diagnostic'):
                if element[1].text == "Record does not exist":
                    record_found = False
                    self.record_status = "Not Found in Worldcat"
                    self.count += 1

            #If the record was found begin parsing the data
            if record_found == True:
                self.record_status = "Bib Info Available"
                for datafield in root.findall(self.namespace+'datafield'):
                    tag = datafield.attrib['tag']
                    # MARC Field 100 (Main Entry - Personal Name) for Author name
                    if tag == '100':
                        for subfield in datafield:
                            if subfield.attrib['code'] == 'a':
                                self.author = subfield.text;
                    # MARC field 245 (Title Statement) for Book Title and more comprehensive author info
                    elif tag == '245':
                        for subfield in datafield:
                            code = subfield.attrib['code']
                            if code == 'a':
                                self.title = subfield.text
                            elif code == 'b':
                                self.title = self.title + ' ' + subfield.text
                            elif code == 'c': #Check subfield c for author, it has more info than Field 100 
                                self.author = subfield.text
                    # MARC Field 250 (Edition Statement) for Book Edition
                    elif tag == '250':
                        for subfield in datafield:
                            if subfield.attrib['code'] == 'a':
                                self.edition = subfield.text
                    # MARC Field 260 (Publication, Distrobution, etc) for publisher and year published
                    # Or MARC Field 264, which has Publication Info as well.  Both use the same subfields
                    elif tag == '260' or tag == '264':
                        for subfield in datafield:
                            code = subfield.attrib['code']
                            if code == 'b':
                                self.publisher = subfield.text
                            elif code == 'c':
                                self.pub_year = subfield.text
                    # MARC Field 020 (International Standard Book Number) for getting price and binding info
                    elif tag == '020':
                        #There are mulitiple subfields with tag 020 so we want to check for the right ISBN to find the right binding
                        correct_isbn = False
                        for subfield in datafield:
                            code = subfield.attrib['code']
                            if code == 'a' and subfield.text == isbn:
                                correct_isbn = True
                            elif code == 'c':
                                self.price = subfield.text
                            elif code == 'q' and correct_isbn == True:
                                #There may be multiple q codes that should be concatenated.
                                self.binding = self.binding + subfield.text
                                #We know we've grabbed the last q code when the text for it ends with a closing parantheses
                                if self.binding[-1] == ')':
                                    correct_isbn = False
                                    continue
                    #With OCLC records price usually appears in Vendor Specific Ordering Data (the 938 field)
                    elif tag == '938':
                        for subfield in datafield:
                            code = subfield.attrib['code']
                            if code == 'a':
                                vendor_hold = subfield.text
                            if code == 'c':
                                self.price = subfield.text
                                self.vendor_name = vendor_hold
                    #OCLC Records don't have price in 365 (where it should be), but adding just in case that changes
                    elif tag == '365':
                        for subfield in datafield:
                            code = subfield.attrib['code']
                            if code == 'b':
                                self.price = subfield.text
                                self.vendor_name = self.publisher

                #OCLC Number is under controlfield element, not datafield, so look for OCLC Number now
                for controlfield in root.findall(self.namespace+'controlfield'):
                    oclc_institution_code = oclc_code.get()
                    tag = controlfield.attrib['tag']
                    if tag == '001':
                        self.oclc_number = controlfield.text
                        #Run API Call for holdings data
                        holdings_url = "https://worldcat.org/webservices/catalog/content/libraries/"+self.oclc_number+"?oclcsymbol="+oclc_institution_code+"&servicelevel=full&wskey="+api_key
                        holdings_api_call = Request(holdings_url, headers={"User-Agent": "Mozilla/5.0"})
                        holdings_raw_data = urlopen(holdings_api_call).read()
                        holdings_root = ET.fromstring(holdings_raw_data)

                        #Assign held_by_institution variable to yes, then check to see if that's true.  If not, reassign variable to no
                        self.held_by_institution = "Yes"
                        for diagnostic in holdings_root.findall(self.message_ns+'diagnostic'):
                            for message in diagnostic:
                                if message.text == "Holding not found":
                                    self.held_by_institution = "No"

                #Remove unwanted characters from OCLC data (unnecessary periods, commas and slashes)
                #Use if statements to avoid index errors when variable is empty
                if self.title != '':
                    self.title = self.title[:-1]
                if self.publisher != '':
                    self.publisher = self.publisher[:-1]
                if self.author != '':
                    self.author = self.author[:-1]
                if self.pub_year != '':
                    if self.pub_year[-1] == '.':
                        self.pub_year = self.pub_year[:-1]

                self.count += 1

            #Write row for current book data
            book_row = [self.record_status, isbn, self.title, self.author, self.publisher, self.pub_year, self.edition, self.price, self.vendor_name, self.binding, self.held_by_institution]
            self.csvwriter.writerow(book_row)

            #Reset Global Variables to avoid lingering data:
            self.record_status = ''
            self.title = ''
            self.author = ''
            self.publisher = ''
            self.pub_year = ''
            self.edition = ''
            self.price = ''
            self.vendor_name = ''
            self.binding = ''
            self.oclc_number = ''
            self.held_by_institution = ''

        #Finished writing, now close the output file
        self.output_file.close()
        return True

    # Check api key makes a preliminary call to OCLC to ensure the key is valid, if not the user gets an error message
    def check_api_key(self, oclc_key):
        isbn = "1111111111111" #Dummy ISBN, results are dependant on api_key only
        api_key = oclc_key.get()
        url = "https://worldcat.org/webservices/catalog/content/isbn/"+isbn+"?servicelevel=full&wskey="+api_key
        api_call = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        try:
            raw_data = urlopen(api_call).read()
            return True
        except URLError as err:
            if err.reason == "Proxy Authentication Required":
                messagebox.showinfo("Error", "API Key is Invalid, Click OK to Stop and Enter a Valid API Key")
            else:
                messagebox.showinfo("ERROR", err.reason)
            return False

    # Get record count to output how many isbns were processed
    def get_record_count(self):
        return str(self.count)


#GUI class for the overlaying user interface
class GUI:
    def __init__(self):
        ######### Window setup #########
        self.window = tk.Tk()
        self.window.geometry("975x550")
        self.window.title("Book Info by ISBN")
        self.window.resizable(0, 0)

        self.header_title = tk.Label(self.window, text="Worldcat Bibliographic API Data to CSV Parser", font=('Arial', 18))
        self.header_title.pack(padx=20, pady=40)

        ######### Setup area for user input #########
        self.entry_frame = tk.Frame(self.window)
        self.entry_frame.columnconfigure(0, weight=1)
        self.entry_frame.columnconfigure(1, weight=1)
        self.entry_frame.columnconfigure(2, weight=1)

        # Row for entering OCLC API Key
        self.key_label = tk.Label(self.entry_frame, text="OCLC API Key:")
        self.key_user_input = tk.StringVar()
        self.key_entry = tk.Entry(self.entry_frame, width=75, textvariable=self.key_user_input)
        self.key_is_saved = tk.IntVar()
        self.key_save = tk.Checkbutton(self.entry_frame, text="Remember Key", variable=self.key_is_saved, command=self.get_save_state)

        # Row for entering OCLC Institutional Code
        self.oclc_code_label = tk.Label(self.entry_frame, text="OCLC Institional Code:")
        self.oclc_code_user_input = tk.StringVar()
        self.oclc_code_entry = tk.Entry(self.entry_frame, width=75, textvariable=self.oclc_code_user_input)
        self.oclc_code_is_saved = tk.IntVar()
        self.oclc_code_save = tk.Checkbutton(self.entry_frame, text="Remember Code", variable=self.oclc_code_is_saved, command=self.get_code_save_state)

        # Row for finding and entering csv to import
        self.file_import_path = tk.Label(self.entry_frame, text="Path to Import File:")
        self.file_import_input = tk.StringVar()
        self.file_import_entry = tk.Entry(self.entry_frame, width=75, textvariable=self.file_import_input)
        self.file_import_browse = tk.Button(self.entry_frame, text="Find File", command=self.open_file)

        # Row for finding directory to save export to
        self.export_path = tk.Label(self.entry_frame, text="Directory to Save Export:")
        self.export_directory_input = tk.StringVar()
        self.export_entry= tk.Entry(self.entry_frame, width=75, textvariable=self.export_directory_input)
        self.export_browse = tk.Button(self.entry_frame, text="Select Directory", command=self.select_directory)

        # Row for naming export file
        self.export_name_label = tk.Label(self.entry_frame, text="Name for Output File:")
        self.export_name_input = tk.StringVar()
        self.export_name_entry = tk.Entry(self.entry_frame, width=75, textvariable=self.export_name_input)

        # Build the grid
        self.key_label.grid(row=1, column=0, sticky=tk.W+tk.E, padx=10, pady=10)
        self.key_entry.grid(row=1, column=1, sticky=tk.W+tk.E, padx=10, pady=10)
        self.key_save.grid(row=1, column=2, sticky=tk.W+tk.E, padx=10, pady=10)
        self.oclc_code_label.grid(row=2, column=0, sticky=tk.W+tk.E, padx=10, pady=10)
        self.oclc_code_entry.grid(row=2, column=1, sticky=tk.W+tk.E, padx=10, pady=10)
        self.oclc_code_save.grid(row=2, column=2, sticky=tk.W+tk.E, padx=10, pady=10)
        self.file_import_path.grid(row=3, column=0, sticky=tk.W+tk.E, padx=10, pady=10)
        self.file_import_entry.grid(row=3, column=1, sticky=tk.W+tk.E, padx=10, pady=10)
        self.file_import_browse.grid(row=3, column=2, sticky=tk.W+tk.E, padx=10, pady=10)
        self.export_path.grid(row=4, column=0, sticky=tk.W+tk.E, padx=10, pady=10)
        self.export_entry.grid(row=4, column=1, sticky=tk.W+tk.E, padx=10, pady=10)
        self.export_browse.grid(row=4, column=2, sticky=tk.W+tk.E, padx=10, pady=10)
        self.export_name_label.grid(row=5, column=0, sticky=tk.W+tk.E, padx=10, pady=10)
        self.export_name_entry.grid(row=5, column=1, sticky=tk.W+tk.E, padx=10, pady=10)

        self.entry_frame.pack()

        # Create a Button to submit user input
        self.submit_button = tk.Button(self.window, text="Submit", command=self.run_submit)
        self.submit_button.pack()

        #  Add a progress bar so the user knows the program is processing
        self.progress_bar = ttk.Progressbar(self.window, orient="horizontal", length=300, mode="indeterminate")
        self.progress_var = tk.StringVar()

        self.status_label_exists = False #this stops the program from creating multiple instances of tkinter status_label

        ######### Create right-click popup menus for copying, cutting, and pasting #########

        # Right click menu for API key entry
        self.context_menu_key_entry = tk.Menu(self.key_entry, tearoff=0)
        self.context_menu_key_entry.add_command(label="Copy", command= self.copy)
        self.context_menu_key_entry.add_separator()
        self.context_menu_key_entry.add_command(label="Cut", command= self.cut)
        self.context_menu_key_entry.add_separator()
        self.context_menu_key_entry.add_command(label="Paste", command= self.paste)
        self.key_entry.bind("<Button - 3>", self.key_context_menu)

        # Right click menu for oclc instituional code entry
        self.context_menu_code_entry = tk.Menu(self.oclc_code_entry, tearoff=0)
        self.context_menu_code_entry.add_command(label="Copy", command= self.copy)
        self.context_menu_code_entry.add_separator()
        self.context_menu_code_entry.add_command(label="Cut", command= self.cut)
        self.context_menu_code_entry.add_separator()
        self.context_menu_code_entry.add_command(label="Paste", command= self.paste)
        self.oclc_code_entry.bind("<Button - 3>", self.code_context_menu)

        # Right click menu for file import entry
        self.context_file_import_entry = tk.Menu(self.file_import_entry, tearoff=0)
        self.context_file_import_entry.add_command(label="Copy", command= self.copy)
        self.context_file_import_entry.add_separator()
        self.context_file_import_entry.add_command(label="Cut", command= self.cut)
        self.context_file_import_entry.add_separator()
        self.context_file_import_entry.add_command(label="Paste", command= self.paste)
        self.file_import_entry.bind("<Button - 3>", self.input_file_context_menu)

        # Right click menu for file export path entry
        self.context_file_export_entry = tk.Menu(self.export_entry, tearoff=0)
        self.context_file_export_entry.add_command(label="Copy", command= self.copy)
        self.context_file_export_entry.add_separator()
        self.context_file_export_entry.add_command(label="Cut", command= self.cut)
        self.context_file_export_entry.add_separator()
        self.context_file_export_entry.add_command(label="Paste", command= self.paste)
        self.export_entry.bind("<Button - 3>", self.export_file_context_menu)

        # Right click menu for export file name entry
        self.context_file_name_entry = tk.Menu(self.export_name_entry, tearoff=0)
        self.context_file_name_entry.add_command(label="Copy", command= self.copy)
        self.context_file_name_entry.add_separator()
        self.context_file_name_entry.add_command(label="Cut", command= self.cut)
        self.context_file_name_entry.add_separator()
        self.context_file_name_entry.add_command(label="Paste", command= self.paste)
        self.export_name_entry.bind("<Button - 3>", self.export_file_name_context_menu)

    ####### GUI Methods #########

    def key_context_menu(self, event):
        self.context_menu_key_entry.tk_popup(event.x_root, event.y_root)
        self.target = 1
    
    def code_context_menu(self, event):
        self.context_menu_code_entry.tk_popup(event.x_root, event.y_root)
        self.target = 2

    def input_file_context_menu(self, event):
        self.context_file_import_entry.tk_popup(event.x_root, event.y_root)
        self.target = 3

    def export_file_context_menu(self, event):
        self.context_file_export_entry.tk_popup(event.x_root, event.y_root)
        self.target = 4

    def export_file_name_context_menu(self, event):
        self.context_file_name_entry.tk_popup(event.x_root, event.y_root)
        self.target = 5


    def copy(self):
        if self.target == 1:
            inp = self.key_entry.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
        elif self.target == 2:
            inp = self.oclc_code_entry.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
        elif self.target == 3:
            inp = self.file_import_entry.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
        elif self.target == 4:
            inp = self.export_entry.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
        elif self.target == 5:
            inp = self.export_name_entry.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)

    def cut(self):
        if self.target == 1:
            inp = self.key_entry.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
            self.key_user_input.set("")
        elif self.target == 2:
            inp = self.oclc_code_entry.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
            self.oclc_code_user_input.set("") 
        elif self.target == 3:
            inp = self.file_import_entry.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
            self.file_import_input.set("") 
        elif self.target == 4:
            inp = self.export_directory_input.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
            self.export_directory_input.set("")
        elif self.target == 5:
            inp = self.export_name_input.get()
            self.window.clipboard_clear()
            self.window.clipboard_append(inp)
            self.export_name_input.set("")

    def paste(self):
        if self.target == 1:
            clipboard = self.window.clipboard_get()
            self.key_user_input.set("")
            self.key_entry.insert('end', clipboard)
        elif self.target == 2:
            clipboard = self.window.clipboard_get()
            self.oclc_code_user_input.set("") 
            self.oclc_code_entry.insert('end', clipboard)
        elif self.target == 3:
            clipboard = self.window.clipboard_get()
            self.file_import_input.set("") 
            self.file_import_entry.insert('end', clipboard)
        elif self.target == 4:
            clipboard = self.window.clipboard_get()
            self.export_directory_input.set("")
            self.export_entry.insert('end', clipboard)
        elif self.target == 5:
            clipboard = self.window.clipboard_get()
            self.export_name_input.set("")
            self.export_name_entry.insert('end', clipboard)

    # File Explorer
    def open_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=(("csv files", "*.csv"),))
        self.file_import_input.set(self.file_path)

    #Directory Selector
    def select_directory(self):
        self.directory_path = filedialog.askdirectory(title="Choose the directory to save to")
        self.export_directory_input.set(self.directory_path)
    
    #Get Key Save State
    def get_save_state(self):
        return self.key_is_saved.get()

    def get_code_save_state(self):
        return self.oclc_code_is_saved.get()

    # Set the Default for the Save API Checkbutton to True
    def set_checkbutton_default_true(self):
        self.key_save.select()

    # Set the Default for the Save OCLC Code Checkbutton to True
    def set_code_checkbutton_default_true(self):
        self.oclc_code_save.select()

    #Start Main Loop
    def start_program_loop(self):
        self.window.mainloop()

    def exit_protocol(self):
        selection = askyesno(title="Close Program", message="Do you wish to close the program ?")
        if selection:
            db = Database()
            if self.get_save_state() == True:
                db.change_api_key(self.get_user_input_key())
                db.change_save_state(True)
            else:
                db.change_save_state(False)
            if self.get_code_save_state() == True:
                db.change_institutional_code(self.get_user_input_oclc_code())
                db.change_code_save_state(True)
            else:
                db.change_code_save_state(False)
            self.window.destroy()

    def end_program(self):   
        self.window.protocol("WM_DELETE_WINDOW", self.exit_protocol)

    #Get the API Key the User Inputs in the Entry Feild
    def get_user_input_key(self):
        return self.key_user_input

    def set_saved_api_key(self, k):
        self.key_user_input.set(k)

    def set_saved_inst_code(self, c):
        self.oclc_code_user_input.set(c)

    # Get the User Input OCLC Code
    def get_user_input_oclc_code(self):
        return self.oclc_code_user_input

    #Get The File Path Selected by the User
    def get_open_file_path(self):
        return self.file_import_input.get()

    #Get Directory Selected by User
    def get_user_directory_input(self):
        return self.export_directory_input.get()

    #Get the User Input for the Output Filename
    def get_user_filename(self):
        if self.export_name_input.get()[-4:] != ".csv":
            self.export_name_input.set(self.export_name_input.get()+".csv")
        return self.export_name_input.get()

    #Get user inputs when submit button is clicked
    def run_submit(self):
        self.submit_button['state'] = DISABLED
        self.user_input_key = self.get_user_input_key()
        self.user_input_oclc_code = self.get_user_input_oclc_code()
        self.user_input_file_path = self.get_open_file_path()
        self.user_input_directory = self.get_user_directory_input()
        self.user_input_file_name = self.get_user_filename()

        if self.user_input_key == "" or \
           self.user_input_oclc_code == "" or \
           self.user_input_file_path == "" or \
           self.user_input_directory == "" or \
           self.user_input_file_name == "":
            self.submit_button['state'] = NORMAL
            messagebox.showinfo("Complete Required Fields", "All Fields Required.  Click OK and Complete Blank Fields")
        else:
            self.parser = Csv_Parser()
            if self.parser.check_api_key(self.user_input_key):
                self.pbar_thread = threading.Thread(target=self.start_progress_bar)
            else:
                self.submit_button['state'] = NORMAL
                return
            if self.parser.open_import_csv(self.user_input_file_path):
                pass
            else:
                self.submit_button['state'] = NORMAL
                self.file_import_input.set("")
                return
            if self.parser.open_export_csv(self.user_input_directory, self.user_input_file_name):
                pass
            else:
                self.submit_button['stat'] = NORMAL
                self.export_directory_input.set("")
                return
           
            self.parser_thread = threading.Thread(target=self.run_parser)
            self.pbar_thread.start()
            self.parser_thread.start()
            

    def start_progress_bar(self):
        self.progress_bar.pack(pady=30)
        self.progress_bar.start(5)
        self.progress_var.set("Processing Please Wait...")
        if self.status_label_exists == False:
            self.status_label = tk.Label(self.window, textvariable = self.progress_var)
            self.status_label.pack()
            self.status_label_exists = True

    def stop_progress_bar(self):
        self.progress_bar.stop()
        self.progress_var.set("Processing Complete.  Total Records Processed: "+self.parser.get_record_count())


    def start_progress_thread(self):
        pthread = threading.Thread(target=self.start_progress_bar)
        pthread.start()

    def run_parser(self):
        while self.parser.get_oclc_data(self.user_input_key, self.user_input_oclc_code) != True:
            pass
        self.stop_progress_bar()
        self.submit_button['state'] = NORMAL
        return True

    def run_parser_thread(self):
        self.parser_thread = threading.Thread(target=self.run_parser)
        if self.parser.check_api_flag() != False:
            self.parser_thread.start()


class Database:
    def __init__(self):
        #Connect to Database and Create Table if it doesn't already exist.
        self.connection = sqlite3.connect('app.db')
        self.cursor = self.connection.cursor()
        self.cursor.execute("""CREATE TABLE IF NOT EXISTS app_data(
            id INTEGER PRIMARY KEY DEFAULT 1,
            save_state BOOLEAN DEFAULT False,
            api_key TEXT DEFAULT 'No Key Saved',
            save_code_state BOOLEAN DEFAULT False,
            oclc_code TEXT DEFAULT 'No Code Saved')""")
        if self.get_save_state() == None:
            self.cursor.execute("INSERT INTO app_data DEFAULT VALUES")
            self.connection.commit()
        self.cursor.close()
        self.connection.close()

    ######### Database Methods #########

    # Get API Key Save State
    def get_save_state(self):
        connection = sqlite3.connect('app.db')
        cursor = connection.cursor()
        state = cursor.execute("SELECT save_state FROM app_data;")
        state = cursor.fetchone()
        cursor.close()
        connection.close()
        return state
    # Change API Key Save State
    def change_save_state(self, state):
        connection = sqlite3.connect('app.db')
        cursor = connection.cursor()
        cursor.execute("UPDATE app_data SET save_state = ?;", (state,))
        connection.commit()
        cursor.close()
        connection.close()

    # Get OCLC Institutional Code Save State
    def get_code_save_state(self):
        connection = sqlite3.connect('app.db')
        cursor = connection.cursor()
        state = cursor.execute("SELECT save_code_state FROM app_data;")
        state = cursor.fetchone()
        cursor.close()
        connection.close()
        return state

    # Change OCLC Institutional Code Save State
    def change_code_save_state(self, state):
        connection = sqlite3.connect('app.db')
        cursor = connection.cursor()
        cursor.execute("UPDATE app_data SET save_code_state = ?;", (state,))
        connection.commit()
        cursor.close()
        connection.close()

    # Get API KEY
    def get_api_key(self):
        connection = sqlite3.connect('app.db')
        cursor = connection.cursor()
        key = cursor.execute("SELECT api_key FROM app_data;")
        key = cursor.fetchone()
        cursor.close()
        connection.close()
        return key[0]

    # Change API KEY
    def change_api_key(self, key):
        connection = sqlite3.connect('app.db')
        cursor = connection.cursor()
        cursor.execute("UPDATE app_data SET api_key = ?;", (key.get(),))
        connection.commit()
        cursor.close()
        connection.close()

    # Get OCLC Institutional Code
    def get_institutional_code(self):
        connection = sqlite3.connect('app.db')
        cursor = connection.cursor()
        key = cursor.execute("SELECT oclc_code FROM app_data;")
        key = cursor.fetchone()
        cursor.close()
        connection.close()
        return key[0]
    # Change API KEY
    def change_institutional_code(self, code):
        connection = sqlite3.connect('app.db')
        cursor = connection.cursor()
        cursor.execute("UPDATE app_data SET oclc_code = ?;", (code.get(),))
        connection.commit()
        cursor.close()
        connection.close()

    
############# MAIN PROGRAM #############
gui = GUI()
database = Database()

#api_key = ""
#file_path = ""
#directory_path = ""
#output_file_name = ""

if database.get_save_state() == (True,):
    gui.set_checkbutton_default_true()
    api_key = database.get_api_key()
    gui.set_saved_api_key(str(api_key))

if database.get_code_save_state() == (True,):
    gui.set_code_checkbutton_default_true()
    oclc_code = database.get_institutional_code()
    gui.set_saved_inst_code(str(oclc_code))

gui.end_program()
gui.start_program_loop()
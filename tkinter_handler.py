import sys

if __name__ == "__main__":
    print("This is not the main module. Do not execute directly.")
    sys.exit()

import os, pyperclip
from tkinter import Tk, Button, filedialog, messagebox, Label, Frame, Canvas, LEFT, SUNKEN
from PIL import Image, ImageTk
from pdf_processing import PdfProcessing
from error_logger import ErrorLogger
from pandas_handler import DataframeHandler
from pcc_handler import PccHandler

class CCWindow:
    def __init__(self,
                 pdf_default_dir: str,
                 excel_file_path: str,
                 excel_file_sheet_name: str,
                 excel_data_first_row: int,
                 pcc_url: str,
                 window_x: int,
                 window_y: int,
                 template_file_path: str,
                 new_diag_button_x: str,
                 code_field_x: str,
                 code_desc_x: str,
                 admis_date_id: str,
                 rank_x: str,
                 clasif_x: str,
                 comm_x: str,
                 confid_x: str,
                 log_file_name: str,
                 FID: str
                 ):
        # --------- PDF settings/data ---------
        self.pdf_default_dir = pdf_default_dir # Default directory to open when user is selecting a PDF to import
        self.pdf_file_dir = None # Directory of PDF user selects
        self.extracted_codes = None # ICD-10 codes extracted from PDF
        self.reset_pdf_text_extracted() # Sets self.pdf_text_extracted to default values of None
        self.import_fails_copy = "" # Text that will be copied to clipboard about code failures

        # --------- Button setup ---------
        self.button_list = [] # List of all buttons in main window (used to disable/enable)
        self.main_buttons_text_col = 0 # Label column for main buttons' instructions in window
        self.main_buttons_col = 1 # Label column for main buttons in window
        self.main_buttons_padx = (30, 10) # Padx for step buttons
        self.main_buttons_pady = 15 # Pady for step buttons
        self.main_button_width = 10 # Width of buttons in step frame
        self.lower_buttons_row = 1 # Label row for lower buttons (copy, current count, exit)
        self.lower_buttons_padx = 25 # Padx for lower buttons (copy, current count, exit)
        self.frame_bg = "#C9C9C9" # Frames' background
        self.line_col = "#7C7C7C" # Line color in steps frame
        self.line_size = 600 # Line width in steps frame
        self.step_1_row = 0 # First row of step 1
        self.step_2_row = 5 # First row of step 2
        self.step_3_row = 7 # First row of step 3
        self.step_4_row = 9 # First row of step 4
        self.lower_frame_pady = (15, 0) # Pady of frame for lower buttons

        # --------- Data used for Excel file in PandasHandler() ---------
        self.excel_file_path = excel_file_path # CM_Codes.xlsx path (just file name)
        self.excel_file_sheet_name = excel_file_sheet_name # Excel sheet codes are in
        self.excel_data_first_row = excel_data_first_row # Data range (cells within the table, not including headers, for writing not reading)
        self.template_file_path = template_file_path # Excel template path (file name). Used to create new CM_Codes.xlsx file when codes are imported.

        self.log_file_name = log_file_name

        # --------- Classes ---------
        """
        self.pcc_handler will be PccHandler() class.
        Declaring here to be able to easily test whether or not the handler has been instantiated (if self.pcc_handler == None).
        Simplest to instantiate later to be able to include loaded DataFrame as argument.
        """
        self.pcc_handler = None

        self.pdf_processor = PdfProcessing()
        self.error_logger = ErrorLogger(self.log_file_name)
        self.dataframe_handler = DataframeHandler(
            self.excel_file_path,
            self.excel_file_sheet_name,
            self.excel_data_first_row,
            self.template_file_path,
            self.log_file_name
        )

        # --------- Setup for PCC interaction ---------
        self.pcc_url = pcc_url
        self.admis_date_id = admis_date_id  # adminDate
        self.FID = FID
        self.cwindow_x = window_x # Chrome window width
        self.cwindow_y = window_y # Chrome window height

        # Xpaths
        self.new_diag_button_x = new_diag_button_x # New Diagnosis button
        self.code_field_x = code_field_x # Field to enter ICD-10 code into
        self.code_desc_x = code_desc_x # Field to read PCC-generated ICD-10 description from
        self.rank_x = rank_x # Rank drop-down menu
        self.clasif_x = clasif_x # Classification drop-down menu
        self.comm_x = comm_x # Comments field
        self.confid_x = confid_x # Confidential checkbox

        # --------- Setup for tkinter window ---------
        self.window_x = 525
        self.window_y = 440

        self.create_win()

    # ------------------ Main Window Setup ------------------
    # Creates main window with widgets
    def create_win(self):
        self.main_window = Tk()

        # Sets window in center of screen
        self.screen_size = (
            self.main_window.winfo_screenwidth(),
            self.main_window.winfo_screenheight()
        )
        self.window_loc_x = self.screen_size[0]/2 - self.window_x/2
        self.window_loc_y = self.screen_size[1]/2 - self.window_y/2

        # Creates window
        self.main_window.geometry(
            f"{self.window_x}x{self.window_y}+"
            f"{int(self.window_loc_x)}+{int(self.window_loc_y)}"
        )
        self.main_window.title("")
        self.main_window.resizable(False, False)

        # Sets window icon (blank for main window, UP icon for children windows)
        self.win_icon = Image.open("IMG_UP_Icon.jpg")
        self.win_icon = ImageTk.PhotoImage(self.win_icon)
        self.main_window.wm_iconphoto(True, self.win_icon) # Sets for children windows

        # self.win_icon = Image.open("IMG_Blank_Icon.jpg")
        # self.win_icon = ImageTk.PhotoImage(self.win_icon)
        # self.main_window.wm_iconphoto(False, self.win_icon) # Sets for parent window only

        # ------------- Canvas for logo (main_window col 0, row 0) -------------
        self.pic_canvas = Canvas(
            self.main_window,
            width = 98,
            height = 366,
            bg = "black"
        )
        self.pic_canvas.grid(column = 0, row = 0)
        self.logo = ImageTk.PhotoImage(Image.open("IMG_MainLogo_NotTransparent_Vertical_Resized2.png"))
        self.pic_canvas.create_image(
            0,
            0,
            anchor = "nw",
            image = self.logo
        )

        # ------------- Steps frame (main_window col 1, row 0) -------------
        self.steps_frame = Frame(
            self.main_window,
            bg = self.frame_bg,
            bd = 1,
            relief = SUNKEN
        )
        self.steps_frame.grid(
            column = self.main_buttons_col,
            columnspan = 3,
            row = self.step_1_row
        )

        # Import/open Excel buttons' instructions (steps_frame col 0, rows 0-3)
        self.place_text("1. Code setup", self.step_1_row, bold = True)
        self.place_text("    Import codes from a PDF file", self.step_1_row + 1)
        self.place_text("    AND/OR", self.step_1_row + 2)
        self.place_text("    Open and edit Excel table of codes", self.step_1_row + 3)

        # Import PDF button (steps_frame col 1, row 0)
        self.import_pdf_button = Button(
            self.steps_frame,
            text ="PDF I\u0332mport",
            width = self.main_button_width,
            command = self.import_pdf
        )
        self.import_pdf_button.grid(
            column = self.main_buttons_col,
            row = self.step_1_row + 1,
            padx = self.main_buttons_padx,
            pady = self.main_buttons_pady,
            sticky = "E"
        )
        self.button_list.append(self.import_pdf_button)
        self.main_window.bind('<Alt-i>', self.import_pdf)

        # Launch CM_Codes.xlsx button (steps_frame col 1, row 3)
        self.open_excel_file_button = Button(
            self.steps_frame,
            text ="O\u0332pen Excel",
            width = self.main_button_width,
            command = self.open_excel_file # Not using lambda function to allow keyboard shortcut to use this
        )
        self.open_excel_file_button.grid(
            column = self.main_buttons_col,
            row = self.step_1_row + 3,
            padx = self.main_buttons_padx,
            pady = self.main_buttons_pady,
            sticky = "E"
        )
        self.button_list.append(self.open_excel_file_button)
        self.main_window.bind('<Alt-o>', self.open_excel_file)

        # Drawing line between steps 1 and 2 (steps_frame col 0, colspan 2, row = 4)
        self.create_line(self.step_1_row + 4)

        # Open Chrome window instructions (steps_frame col 0, row 5)
        self.place_text(
            "2. Open new Chrome window",
            self.step_2_row,
            bold = True
        )

        # Open Chrome window (steps_frame col 1, row 5)
        self.open_chrome_button = Button(
            self.steps_frame,
            text ="Ch\u0332rome",
            width = self.main_button_width,
            command = self.open_chrome
        )
        self.open_chrome_button.grid(
            column = self.main_buttons_col,
            row = self.step_2_row,
            padx = self.main_buttons_padx,
            pady = self.main_buttons_pady,
            sticky = "E"
        )
        self.button_list.append(self.open_chrome_button)
        self.main_window.bind('<Alt-h>', self.open_chrome)

        # Drawing line between steps 2 and 3 (steps_frame col 0, colspan 2, row 6)
        self.create_line(self.step_2_row + 1)

        # Log in/navigate to profile instructions (steps_frame col 0, row 7)
        self.place_text(
            "3. Log into PCC and navigate to the \n"\
                "    Med Diag tab of the patient's profile",
            self.step_3_row,
            bold = True,
            extra_y = True
        )

        # Drawing line between steps 3 and 4 (steps_frame col 0, colspan 2, row 8)
        self.create_line(self.step_3_row + 1)

        # Enter codes instructions (steps_frame col 0, row 9)
        self.place_text(
            "4. Begin the import process\n"\
                "    on the current profile page",
            self.step_4_row,
            bold = True
        )

        # Enter codes into PCC button (steps_frame col 1, row 9)
        self.enter_into_pcc_button = Button(
            self.steps_frame,
            text ="PCC E\u0332ntry",
            width = self.main_button_width,
            command = self.enter_into_pcc
        )
        self.enter_into_pcc_button.grid(
            column = self.main_buttons_col,
            row = self.step_4_row,
            padx = self.main_buttons_padx,
            pady = self.main_buttons_pady,
            sticky = "E"
        )
        self.button_list.append(self.enter_into_pcc_button)
        self.main_window.bind('<Alt-e>', self.enter_into_pcc)

        # ------------- Lower buttons label (main_window col 0, colspan 4, row 1) -------------
        self.lower_buttons_label = Label(self.main_window)
        self.lower_buttons_label.grid(
            column = 0,
            columnspan = 4,
            row = self.lower_buttons_row,
            pady = self.lower_frame_pady
        )

        # Count codes in current Excel file (lower_buttons_label column 0 row 0)
        # (Excel file can be open for openpyxl to read.)
        self.count_curr_codes_button = Button(
            self.lower_buttons_label,
            text = "Cu\u0332rrent Excel code count",
            command = self.code_count
        )
        self.count_curr_codes_button.grid(
            column = 0,
            row = 0,
            padx = self.lower_buttons_padx
        )
        self.button_list.append(self.count_curr_codes_button)
        self.main_window.bind('<Alt-u>', self.code_count)

        # Copy most recent list of PCC import failures (lower_buttons_label column 1 colspan 2 row 0)
        # Not added to self.button_list; this button is treated uniquely so it is only enabled when a fail list exists.
        self.copy_import_fails_button = Button(
            self.lower_buttons_label,
            text = "C\u0332opy most recent import fail info",
            command = self.copy_fails_to_clipb,
            state = "disabled"
        )
        self.copy_import_fails_button.grid(
            column = 1,
            columnspan = 2,
            row = 0,
            padx = self.lower_buttons_padx
        )
        self.main_window.bind('<Alt-c>', self.copy_fails_to_clipb)

        # Exit button  (lower_buttons_label column 3 row 0)
        self.exit_button = Button(
            self.lower_buttons_label,
            text = "Ex\u0332it",
            command = self.close_out
        )
        self.exit_button.grid(
            column = 3,
            row = 0,
            padx = self.lower_buttons_padx
        )
        self.button_list.append(self.exit_button)
        self.main_window.bind('<Alt-x>', self.close_out)

        self.main_window.mainloop()

    # Places instructions text in step frame
    # Receives text to place, row to placei t on, whether it's bold or not, and whether it needs extra y-padding
    def place_text(self,
                   text: str,
                   row: int,
                   bold: bool = False,
                   extra_y: bool = False
                   ):
        if bold == False: font = ("Segoe UI", 9)
        else: font = ("Segoe UI bold", 9)
        if extra_y == True: pady = 10
        else: pady = 0

        self.label_text = Label(
            self.steps_frame,
            text = text,
            bg = self.frame_bg,
            justify = LEFT,
            font = font,
            pady = pady
        )
        self.label_text.grid(
            column = self.main_buttons_text_col,
            row = row,
            sticky = "W"
        )

    # Draws line in steps frame on designated row
    def create_line(self, row_num: int):
        self.line_canvas = Canvas(
            self.steps_frame,
            bg = self.frame_bg,
            height = 1,
            width = self.line_size / 2,
            bd = 0,
            highlightthickness = 0
        )
        self.line_canvas.grid(
            column = self.main_buttons_text_col,
            columnspan = 2,
            row = row_num,
            pady = 10
        )
        self.line_canvas.create_line(
            0, # X0
            0, # Y0
            0, # X1
            self.line_size, # Y1
            width = self.line_size,
            fill = self.line_col
        )

    # Enables buttons after processing (e.g., regex for large PDFs)
    def enable_buttons(self):
        for button in self.button_list:
            button["state"] = "normal"
        if self.import_fails_copy != "": self.copy_import_fails_button["state"] = "normal"

    # Disables buttons during processing (e.g., regex for large PDFs)
    def disable_buttons(self):
        for button in self.button_list:
            button["state"] = "disabled"
        self.copy_import_fails_button["state"] = "disabled"

    # Safely closing out the window
    def close_out(self, e = None):
        if not messagebox.askyesno("Confirm exit", "Do you want to exit CodeM UP?"): return
        self.error_logger.manage_file_size() # Keeps error log capped at 50
        if self.pcc_handler != None: self.pcc_handler.webdriver.close_out()
        self.main_window.quit()

    # ------------------ Excel and PDF management ------------------
    # Imports PDF and scans for ICD-10 codes
    def import_pdf(self, e = None):
        # Ensures that Excel file is closed (cannot write data to the file if it's open)
        if self.pdf_processor.is_excel_file_open(self.excel_file_path):
            messagebox.showerror(
                "Excel file open",
                "Please close the Excel codes file before importing a new set of codes."
            )
            return

        # Ensuring that user wants to proceed and clear out existing codes
        if not messagebox.askyesno(
                title = "Clear codes",
                message = "Importing a new PDF will clear any codes previously imported.\n\n"
                    "Do you want to proceed?"
        ): return

        # Getting file path from user
        self.pdf_file_dir = filedialog.askopenfilename(
            title = "Select PDF",
            initialdir = self.pdf_default_dir,
            filetypes = [("PDFs", "*.pdf")]
        )

        if self.pdf_file_dir == "": return # No file selected by user

        # Disables buttons (large PDFs may require several seconds to process)
        self.import_pdf_button.config(text = "IMPORTING")
        self.disable_buttons()
        self.main_window.update() # Shows buttons as disabled

        # Imports all PDF pages' text
        # Receives a three-item tuple return
        # tup[0] = either a list of all pages' text OR False if an error occured
        # tup[1] = either None (successful text extraction) OR a string message to user about error
        # tup[2] = either None (successful text extractiOn) OR exception captured by try/except
        self.pdf_text_extracted = self.pdf_processor.import_pdf(self.pdf_file_dir)

        if self.pdf_text_extracted[0] == False: # Error encountered in self.PdfProcessing
            # Logs error
            if self.pdf_text_extracted[2] != False: # False indicates an error to the user only (no codes found, etc.)
                self.error_logger.log_error(
                    "Attempted to extract text from PDF",
                    self.pdf_text_extracted[2]
                )

            # Captures error message before resetting self.pdf_text_extracted
            # (Reset to ensure no old data is retained that could cause errors with later processes)
            err_msg = self.pdf_text_extracted[1]
            self.reset_pdf_text_extracted()

            # Informs user of error
            messagebox.showerror("Error", err_msg)
            self.enable_buttons()
            return

        # Gets all codes from extracted text
        # self.pdf_text_extracted[0] is itself a list of text from each page
        self.extracted_codes = self.pdf_processor.apply_regex(self.pdf_text_extracted[0])

        # Creates dataframe from extracted codes
        self.dataframe_handler.create_df(self.extracted_codes)
        response = self.dataframe_handler.save_codes_to_excel()
        if response != None: # None = no errors. If errors, user-friendly text description of error is returned.
            messagebox.showerror("Error encountered", response)
            self.enable_buttons()
            return

        # Notifies user of how many codes were found
        self.import_pdf_button.config(text="PDF I\u0332mport")
        num_codes = len(self.extracted_codes)
        if num_codes == 1: code_sp = "code"
        else: code_sp = "codes"
        messagebox.showinfo("Codes found", f"{len(self.extracted_codes)} {code_sp} were found.")

        # Turns buttons back on after PDF data extraction completes
        self.enable_buttons()

    # Sets all 3 elements of self.pdf_text_extracted to None.
    # Easy way to make sure variable is empty and results from a previous PDF doesn't interfere with a new process.
    def reset_pdf_text_extracted(self): self.pdf_text_extracted = [None, None, None]

    # Opens Excel file
    # Method not lambda function so that keyboard shortcut will work
    def open_excel_file(self, e = None):
        os.system(f"start EXCEL.EXE {self.excel_file_path}")

    # ------------------ User Feedback ------------------
    # Copies all fails to clipboard if user requests it (at end of upload or anytime afterward
    def copy_fails_to_clipb(self, e = None):
        if self.import_fails_copy == "": return
        pyperclip.copy(self.import_fails_copy)
        messagebox.showinfo(
            "Code fails copied",
            "All code fail data have been copied."
        )

    # Sets up text to display to user regarding failed codes/failed code details
    # Output varies if it's displayed to the user in a messagebox of if copied to clipboard.
    def code_text_setup(self,
                        delimiter: str,
                        text_to_user_codes: str = "",
                        text_to_user_details: str = ""
                        ):
        list_failed_codes = ""
        if self.pcc_handler.failed_to_enter_code != []:
            list_failed_codes += f"{text_to_user_codes}\n"
            list_failed_codes += f"{delimiter}".join(self.pcc_handler.failed_to_enter_code)
            list_failed_codes += "\n"
        if self.pcc_handler.failed_to_enter_other != []:
            list_failed_codes += f"{text_to_user_details}\n"
            list_failed_codes += f"{delimiter}".join(self.pcc_handler.failed_to_enter_other)
            list_failed_codes += "\n"
        return list_failed_codes

    # Counts how many codes are in the Excel file. Excel file can be open for openpyxl to read this.
    def code_count(self, e = None):
        self.dataframe_handler.read_codes_from_excel()
        count = len(
            self.dataframe_handler.df_excel_import_codes[
                self.dataframe_handler.df_excel_import_codes[
                    self.dataframe_handler.header_list[0]
                ] != ""
            ]
        )
        if count == 1: row_sp = "row"
        else: row_sp = "rows"
        messagebox.showinfo(
            "Code count",
            f"The Excel file currently has {count} {row_sp} to iterate through."
        )

    # ------------------ Chrome Management ------------------
    # Opens Chrome window
    # Instantiates PccHandler class (which instantiates Webdriver class from WebdriverFramework.py)
    def open_chrome(self, e = None):
        self.disable_buttons()
        self.pcc_handler = PccHandler(
            self.pcc_url,
            self.admis_date_id,
            self.new_diag_button_x,
            self.code_field_x,
            self.code_desc_x,
            self.rank_x,
            self.clasif_x,
            self.comm_x,
            self.confid_x,
            self.log_file_name,
            self.FID,
            self.cwindow_x,
            self.cwindow_y
        )
        self.pcc_handler.open_new_window()
        self.enable_buttons()

    # Uploads data into PCC
    def enter_into_pcc(self, e = None):
        self.disable_buttons()

        # self.pcc_handler set to None at CCWindow instantiation. Set to PccHandler() class when Chrome is opened.
        if self.pcc_handler == None:
            messagebox.showerror(
                "Chrome not started",
                "Chrome has not been started yet. Please start Chrome and navigate to the relevant patient's Med Diag tab before attempting to import codes into PCC."
            )
            self.enable_buttons()
            return

        # Reads back in the data as a dataframe from the Excel file
        response = self.dataframe_handler.read_codes_from_excel()
        if response != None: # None = successful, returned string = error message to user
            messagebox.showerror("Error encountered", response)
            self.enable_buttons()
            return

        # Confirms that user wants to continue
        if not messagebox.askyesno(
                "Ready to enter into PCC?",
                "Are you ready to enter the codes in Excel into PCC?\n\n"
                    "Make sure you are viewing the patient's Med Diag tab before you proceed!\n\n"
                    "You cannot stop this process until it completes on its own."
        ):
            self.enable_buttons()
            return

        # None returned means no errors to prepend to output to user (e.g., error like Chrome not detected/probably closed by user)
        upload_response = self.pcc_handler.upload_all_codes(
            self.dataframe_handler.df_excel_import_codes,
            self.dataframe_handler.header_list
        )
        if upload_response == None: upload_response = ""
        elif upload_response == False: # Error message to user already delivered. Just needs to exit method.
            self.enable_buttons()
            return

        # Informs user that import is complete
        completed_user_msg = "PCC code import process finished." + upload_response # Error info (e.g., user closed Chrome)
        # If any codes failed or code details failed, informs user and gives the option to copy these to the Clipboard
        # Creates list of unique failed codes
        self.pcc_handler.failed_to_enter_code = \
            [ele for i, ele in enumerate(self.pcc_handler.failed_to_enter_code)
             if ele != "" and
             self.pcc_handler.failed_to_enter_code[i:].count(ele) == 1]
        # Creates list of unique failed additions of code details
        self.pcc_handler.failed_to_enter_other = \
            [ele for i, ele in enumerate(self.pcc_handler.failed_to_enter_other)
             if ele != "" and
             self.pcc_handler.failed_to_enter_other[i:].count(ele) == 1]
        # If there are errors, creates user-friendly text output
        if self.pcc_handler.failed_to_enter_code != [] or self.pcc_handler.failed_to_enter_other != []:
            completed_user_msg += "\n"
            # Sets up text in window to user
            completed_user_msg += self.code_text_setup(
                delimiter = ", ",
                text_to_user_codes = "\nThe below list of codes were not entered into PCC. Please manually enter them or investigate why PCC did not accept them.\n",
                text_to_user_details = "\nThe below list of code details were not entered into PCC for individual codes. Please manually enter these.\n"
            )
            completed_user_msg += "\nWould you like to copy this information to your Windows clipboard?"

            # Sets up text the user can copy
            self.import_fails_copy = self.code_text_setup(
                delimiter="\n",
                text_to_user_codes="Codes not entered:\n",
                text_to_user_details="\nDetails not entered:\n"
            )

            if messagebox.askyesno("Import complete", completed_user_msg): self.copy_fails_to_clipb()
        else: # No errors to report
            # Sets text to copy to empty string (in case left over from another import) and disables copy button
            self.import_fails_copy = ""
            self.copy_import_fails_button["state"] = "disabled"
            messagebox.showinfo("Import complete.\n\nNo errors encountered.", completed_user_msg)

        self.enable_buttons()
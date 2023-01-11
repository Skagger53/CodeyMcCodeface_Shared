import sys

if __name__ == "__main__":
    print("This is not the main module. Do not execute directly.")
    sys.exit()

import datetime
import time
from time import sleep
from tkinter import messagebox
import pandas.core.frame
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import selenium.common # Used for error detection when switching windows within methods
from WebdriverFramework import WebdriverMain
from error_logger import ErrorLogger
from DataValidation import DataValidation

class PccHandler:
    def __init__(self,
                 pcc_url: str,
                 admis_date_id: str,
                 new_diag_button_x: str,
                 code_field_x: str,
                 code_desc_x: str,
                 rank_x: str,
                 clasif_x: str,
                 comm_x: str,
                 confid_x: str,
                 log_file_name: str,
                 FID: str,
                 window_x: int,
                 window_y: int
                 ):

        # Xpaths
        self.new_diag_button_x = new_diag_button_x # New Diagnosis button
        self.code_field_x = code_field_x # Field to enter ICD-10 code
        self.code_desc_x = code_desc_x # Field with code description (used to verify that a code has been found by PCC)
        self.rank_x = rank_x # Drop-down for Rank
        self.clasif_x = clasif_x # Drop-down for Classification
        self.comm_x = comm_x # Comments field
        self.confid_x = confid_x # Confidential checkbox

        # Classes
        self.DataValidation = DataValidation()
        self.ErrorLogger = ErrorLogger(log_file_name)

        # Misc assignments
        self.df_excel_import_codes = None # Dataframe of codes from Excel file
        self.pcc_url = pcc_url # URL to log into PCC
        self.admis_date_id = admis_date_id  # ID for admission date link
        self.FID = FID
        self.fail_to_man_win = (
            "Window management error",
            "Failed to interact with Chrome windows.\n\n"
                "The PCC Import process will terminate. Wait a few seconds and try again."
        ) # Failure message if software cannot reduce number of windows. Very unlikely to happen. Terminates import process if this happens.
        self.fail_to_verify = "This software is only licensed for use at the Estates at Chateau.\n\n" \
                              "If you would like to use this software for another facility, contact Matt Skaggs (matt.reword@gmail.com) to discuss licensing."

        self.failed_to_enter_code = [] # List of codes that failed to enter
        self.failed_to_enter_other = [] # List of other items that failed to enter

        self.window_x = window_x # Width for Chrome window
        self.window_y = window_y # Height for Chrome window

    # Starts a new webdriver and navigates to PCC page
    def open_new_window(self):
        self.webdriver = WebdriverMain(
            suppress_all_msgs = True,
            suppress_confirmation = True,
            window_x = self.window_x,
            window_y = self.window_y
        )
        sleep(2) # New installations can encounter an error
        self.webdriver.get_url(self.webdriver.main_win_handle, self.pcc_url)

    # --------------------- Entering Codes ---------------------
    # Enters all code data into PCC
    # This calls all below methods for entering codes
    def upload_all_codes(self,
                         df_excel_import_codes: pandas.core.frame.DataFrame,
                         excel_headings: tuple
                         ):

        # Tests if Chrome is open (user could have opened then closed it)
        # Already tested to see that Chrome was instantiated once.
        try: self.webdriver.driver.switch_to.window(self.webdriver.driver.window_handles[0])
        except: return "\n\nNo Chrome windows/tabs detected. Is Chrome open?"

        # Reduces all open windows/tabs to 1 if more than one is open. Informs user of this and gives them the chance to close out other windows/tabs.
        num_windows = self.num_windows()
        if num_windows == False: # Failed to count the windows. Could be from too many attempts too quickly.
            messagebox.showerror(self.fail_to_man_win[0], self.fail_to_man_win[1])
            return False
        elif self.num_windows() > 1:
            messagebox.showinfo(
                "Close windows/tabs",
                "Chrome has more than one window/tab open.\n\n"
                    "All windows/tabs but one will now be closed. (You may want to manually close all windows/tabs now except the one open at the patient's Med Diag tab.)"
            )

            if self.reduce_win_to_one() == False:
                messagebox.showerror(self.fail_to_man_win[0], self.fail_to_man_win[1])
                return False

        try: webd_ele = self.webdriver.find_ele(self.webdriver.main_win_handle, "id", self.FID, "Facility name")
        except Exception as find_webd_ele_e:
            self.ErrorLogger.log_error("Attempted to find webdriver element", find_webd_ele_e)
            messagebox.showerror(
                "Could not verify PCC webpage",
                "Could not verify that Chrome is at PCC.\n\n"
                        "Are you logged into PCC and at the patient's Med Diag tab?"
            )
            return False
        if self.check(webd_ele) == False: return False

        # Checks that New Diagnosis button can be found at all.
        new_diag_button = self.webdriver.find_ele(self.webdriver.main_win_handle, "xpath", self.new_diag_button_x, "New Diagnosis button")
        if new_diag_button == False:
            messagebox.showerror(
                "Could not verify PCC webpage",
                "Could not verify that Chrome is at the Med Diag tab.\n\n"
                    "Are you logged into PCC and at the Med Diag tab?"
            )
            return False

        self.df_excel_import_codes = df_excel_import_codes
        self.excel_headings = excel_headings

        # The below are set here so it resets each time the user enters the codes into PCC (Excel code data may have been manually changed, new PDF may have been imported, etc.)
        self.code_field_obj = None # Webdriver object for field to enter ICD-10 code
        self.admission_date = True # True = admission date available for patient; False = not available (no admit census line)
        self.date_obj_list = None # List of webdriver objects for elements with "pccDateField" ID (admit and resolved date fields)
        self.rank_obj = None # Webdriver object for Rank drop-down
        self.classif_obj = None # Webdriver object for Classification drop-down
        self.comm_obj = None # Webdrive obejct for Comments field
        self.save_buttons = None # List of webdriver objects for elements with "pccButton saveButtons" ID ([1] is Save and [2] is Save & New)
        self.failed_to_enter_code = [] # List of codes that failed to enter
        self.failed_to_enter_other = [] # List of other items that failed to enter

        df_header_list = self.df_excel_import_codes.columns.values
        cons_failed_iter = 0

        # Iterates through all code data
        for i in range(0, len(self.df_excel_import_codes.index)):
            # Used in below for col in df_header_list loop to break out of an iteration (but is reset in each outside loop)
            skip_col_iteration = False
            current_code = self.df_excel_import_codes.loc[i, df_header_list[0]].strip() # ICD-10 code for this iteration
            if current_code == "": continue # Skips empty strings

            # Closes most recently opened window(s) until there is only one window open
            # Effectively resets back to the main PCC page each time
            if self.reduce_win_to_one() == False:
                messagebox.showerror(self.fail_to_man_win[0], self.fail_to_man_win[1])
                for failed_codes in range(i, len(self.df_excel_import_codes.index)): # Logs remaining codes as failed
                    self.failed_to_enter_code.append(self.df_excel_import_codes.loc[failed_codes, df_header_list[0]].strip())
                return "\n\nCould not manage Chrome windows, so the PCC import process was terminated.\n\nPlease wait a few seconds and try again."

            if i != 0: sleep(3) # No need to wait on very first iteration
            self.webdriver.driver.switch_to.window(self.webdriver.driver.window_handles[0])
            self.webdriver.main_win_handle = self.webdriver.driver.window_handles[0]

            # Clicks New Diagnosis button
            # If software can't find the button twice in a row, terminates iteration, logs all remaining codes as failed, and informs user.
            click_new_diag = self.webdriver.click_ele(self.webdriver.main_win_handle, new_diag_button, "New Diagnosis button")
            if click_new_diag == False:
                self.webdriver.driver.refresh()
                sleep(.25)
                click_new_diag = self.webdriver.find_click(self.webdriver.main_win_handle, "xpath", self.new_diag_button_x, "New Diagnosis button")
                if click_new_diag == False: # Failed to find button
                    self.failed_to_enter_code.append(current_code)
                    cons_failed_iter += 1
                    if cons_failed_iter >= 2: # 2 failed attempts in a row ends entering codes into PCC
                        for failed_codes in range(i, len(self.df_excel_import_codes.index)): # Logs remaining codes as failed
                            self.failed_to_enter_code.append(self.df_excel_import_codes.loc[failed_codes, df_header_list[0]].strip())
                        return "\n\nCould not find the New Diagnosis button, so the PCC import process was terminated.\n\nIs Chrome on the correct webpage?"
                    continue

            cons_failed_iter = 0 # If successfully finds button, resets counter of failed attempts

            sleep(.25) # Gives enough time for window to open
            self.get_diag_win() # Properly sets self.diag_win

            # Iterates through all details for each code
            for col in df_header_list:
                # May be set to skip if code fails to enter (no need to attempt to enter that code's details)
                if skip_col_iteration == True: break
                data_value = self.df_excel_import_codes.loc[i, col].strip()

                """ Skip if blank for all except Rank and Classification.
                    According to Regional Nursing Consultant Amber Baudler, unless otherwise specified,
                    Rank should be "Other" and Classifiation should always be "Admission". """
                if data_value == "":
                    if col == self.excel_headings[2]: data_value = "Other" # Rank
                    elif col == self.excel_headings[3]: data_value = "Admission" # Classification
                    else: continue

                # According to Regional Nurse Consultant Amber Baudler, always use admit date for the Date field
                # (though this will enter the current date if there is no admit date in the census line)
                self.enter_adm_date(current_code)

                if col == self.excel_headings[0]: # 'ICD-10 Code'
                    if self.enter_diag_code(data_value) == False:
                        skip_col_iteration = True
                        continue
                elif col == self.excel_headings[1]: # "Resolved Date"
                    self.enter_resolved_date(data_value, current_code)
                elif col == self.excel_headings[2]: # "Rank"
                    self.enter_rank(data_value, current_code)
                elif col == self.excel_headings[3]: # "Classification"
                    self.enter_classif(current_code)
                elif col == self.excel_headings[4]: # "Comments"
                    self.enter_comments(data_value, current_code)
                elif col == self.excel_headings[5]: # "Confidential?"
                    self.select_confidential(data_value, current_code)

            self.click_save(current_code)

            sleep(.2)

            # Checks for an error in entering a code (e.g., code already exists for this patient)
            self.webdriver.driver.switch_to.window(self.webdriver.driver.window_handles[-1])
            try: self.webdriver.driver.find_element(By.ID, "pccError")
            except: pass
            else: self.failed_to_enter_code.append(current_code) # Adds this code to list of failed codes

        self.reduce_win_to_one() # Full import process complete. No need to check if window reduction failed.

    # Enters a diagnosis code
    def enter_diag_code(self, code: str):
        code_field_msg = "Code text box in New Diagnosis window" # Fail message passed to webdriver find_ele method

        # if self.code_field_obj == None:
        self.code_field_obj = self.webdriver.find_ele(self.diagn_win, "xpath", self.code_field_x, code_field_msg)

        if self.code_field_obj == False: # Could not find the field at all to begin searching for code
            self.failed_to_enter_code.append(code)
            return False

        enter_code = self.webdriver.enter_text_ele(self.diagn_win, self.code_field_obj, code, code_field_msg) # Enters code into found Code field
        if enter_code == False:
            self.failed_to_enter_code.append(code)
            return False

        # Tabbing to the next field usually opens a new window either to select specific code or to inform user
        # that the code is not specific enough -- the latter must be prevented, as the window loads infinitely without
        # manual input.
        self.code_field_obj.send_keys(Keys.TAB)

        # The below is a solution to the most challenging part of this project.
        # This attempts to keep PCC from opening a window with a prompt on top of an infinitely loading webpage;
        # Selenium cannot interact with a prompt on a webpage that's still loading, so the entire import process stops if this happens.
        # The software would hang forever.
        # This looks for 3 windows, which means either (1) a window has appeared requiring the user to select the All Diagnoses link"
        # (frequently required) or (2) the window is appearing that will have the infinitely loading page with
        # the prompt on it.
        # The below process runs for 0.2 seconds and rapidly checks for three windows being open at the same time.
        # If there are three, it attempts to click the All Diagnoses link. If it cannot find the link,
        # presumably the window is about to load infinitely with the prompt on top,
        # and it closes that window before the prompt can even appear.
        timer_start = time.time()
        while timer_start + .2 > time.time():
            num_windows = self.num_windows()
            if num_windows == False: self.close_most_rec_win()
            elif self.num_windows() == 3:
                # Tries to get the All Diagnosis link. If an ICD-10 code is in PCC, this will ensure it is found and assigned
                # This may fail if PCC is going to display an alert error (which cannot be closed out because the window loads infinitely without manual user input)
                try:
                    self.webdriver.driver.switch_to.window(self.webdriver.driver.window_handles[-1]) # Switches to new window
                    all_diag_link = self.webdriver.driver.find_elements(By.CLASS_NAME, "viewFilter")[0] # Tries to get the link
                    self.webdriver.click_ele(self.webdriver.driver.window_handles[-1], all_diag_link, "All Diagnoses link") # Tries to click the link
                except: # Try above failed. Presumably there is no link to click and the window is about to load infinitely.
                    try: self.webdriver.driver.switch_to.alert.accept() # Attempts to dismiss alert (this has never worked in testing)
                    except: pass
                    finally:
                        num_windows = self.num_windows()
                        if num_windows == 3 or num_windows == False: self.close_most_rec_win() # Closes out the most recent window.
                break

        # Obtains PCC-generated code description. Returns False if failed to get code (presumably the window is not open)
        code_desc = self.get_desc()

        if code_desc == False: # Failed to find description at all. Failed to enter this code.
            self.failed_to_enter_code.append(code)
            return False

        if code_desc == "": # Found the description field, but PCC did not generate a description
            self.failed_to_enter_code.append(code) # Appends code as failure (PCC didn't recognize code)
            return False

    # Fills in the Date field (uses admission date link)
    def enter_adm_date(self, code: str):
        fail_adate_txt = "Admission date link" # Used if entering this data fails

        self.admission_date_obj = self.webdriver.find_click(self.diagn_win, "id", self.admis_date_id, fail_adate_txt)
        if self.admission_date_obj == False:
            print(f"Failed to find {fail_adate_txt}")
            self.fail_other_log(code, "Admission date", "Date field")
            return

        # If an alert appears, this dismisses it
        self.dismiss_alert()

    # Fills in the Resolved Date field
    def enter_resolved_date(self, date: str, code: str):
        fail_rdate_txt = "Resolved Date field" # Used if entering this data fails

        # Attempts to get date not time portion (Pandas imports as datetime object with 00:00:00 as time, but this only needs the date portion as string)
        try: date = date.split()[0]
        except: pass

        # Validates that the imported text is a date
        validate_date = self.DataValidation.validate_user_input_date(date)
        if validate_date == False:
            self.fail_other_log(code, date, fail_rdate_txt)
            return
        date = datetime.datetime.strftime(validate_date, "%m/%d/%Y") # Formats date the way PCC expects

        # This obtains list of element objects of class "pccDateField" for future reference. Desired element is [1] ([0] is admission date).
        # Failure records failure and returns.
        self.date_obj_list = self.webdriver.driver.find_elements(By.CLASS_NAME, "pccDateField")
        try:
            self.date_obj_list = self.webdriver.driver.find_elements(By.CLASS_NAME, "pccDateField")
        except Exception as find_resolved_date_e:
            self.fail_other_log(code, date, fail_rdate_txt)
            return
        else:
            # Checks if the validation date is less than the admit date. PCC does not allow this.
            admit_date_obj = datetime.datetime.strptime(self.date_obj_list[0].get_attribute("value"), "%m/%d/%Y")
            try:
                admit_date_obj = datetime.datetime.strptime(self.date_obj_list[0].get_attribute("value"), "%m/%d/%Y")
            # Can't parse date (probably user error, since Excel is configured to format dates properly)
            except Exception as parse_admit_date_e:
                self.fail_other_log(code, date, fail_rdate_txt)
                return
            else:
                if validate_date < admit_date_obj:
                    self.fail_other_log(code, date, fail_rdate_txt)
                    return

        # Enters the date as string into proper field
        if self.webdriver.enter_text_ele(self.diagn_win, self.date_obj_list[1], date, "Resolved Date field") == False:
            self.fail_other_log(code, date, fail_rdate_txt)

    # Fills in the rank
    # Per Regional Nurse Consultant Amber Baudler, software will always select Other (unless otherwise specified)
    def enter_rank(self, rank: str, code: str):
        fail_rank_txt = "Rank drop-down menu" # Used if entering this data fails

        self.rank_obj = self.webdriver.find_ele(self.diagn_win, "xpath", self.rank_x, fail_rank_txt)
        if self.rank_obj == False:
            self.fail_other_log(code, rank, fail_rank_txt)
            return

        # Enters the rank into Rank drop-down
        if self.webdriver.enter_text_ele(self.diagn_win, self.rank_obj, rank, "Rank drop-down menu") == False:
            self.fail_other_log(code, rank, fail_rank_txt)

    # Enter Classification
    # Per Regional Nurse Consultant Amber Baudler, software will always select Admission (unless otherwise specified)
    def enter_classif(self, code: str):
        fail_classif_txt = "Classification drop-down" # Used if entering this data fails

        self.classif_obj = self.webdriver.find_ele(self.diagn_win, "xpath", self.clasif_x, fail_classif_txt)
        if self.classif_obj == False:
            print(f"Failed to find {fail_classif_txt}")
            self.fail_other_log(code, "Admission", fail_classif_txt)
            return

        # Attempts to enter "Admission" as selection
        if self.webdriver.enter_text_ele(self.diagn_win, self.classif_obj, "Admission", fail_classif_txt) == False:
            self.fail_other_log(code, "Admission", fail_classif_txt)

    # Enters comments
    def enter_comments(self, comment: str, code: str):
        fail_comm_txt = "Comments field" # Used if entering this data fails

        self.comm_obj = self.webdriver.find_ele(self.diagn_win, "xpath", self.comm_x, fail_comm_txt)
        if self.comm_obj == False:
            print(f"Failed to find {fail_comm_txt}")
            self.fail_other_log(code, comment, fail_comm_txt)
            return

        # Attempts to enter "Admission" as selection
        if self.webdriver.enter_text_ele(self.diagn_win, self.comm_obj, comment, fail_comm_txt) == False:
            self.fail_other_log(code, comment, fail_comm_txt)

    # Enables the Confidential? check box if necessary
    def select_confidential(self, data_value: str, code: str):
        if data_value[0].lower() != "y": return # Anything but a "yes" is returned. Already ensured len() > 0.

        if self.webdriver.find_click(self.diagn_win, "xpath", self.confid_x, "Confidential checkbox") == False:
            self.fail_other_log(code, "check", "'Confidential' checkbox")

    # Tries to get text from description field. Returns text found or False if fails.
    def get_desc(self):
        sleep(.5)
        try: self.webdriver.driver.switch_to.window(self.diagn_win)
        except: return False # self.diagn_win must be closed
        code_desc = self.webdriver.find_ele(self.diagn_win, "xpath", self.code_desc_x, "Code Description")
        if code_desc == False: return False
        return code_desc.get_attribute("value")

    # Click Save
    def click_save(self, code: str):
        fail_save_txt = "Save button"
        num_wins = len(self.webdriver.driver.window_handles)

        try: self.save_buttons = self.webdriver.driver.find_elements(By.CLASS_NAME, "saveButtons")
        except Exception as find_save_e:
            self.failed_to_enter_code.append(code)
            return False

        try:
            self.webdriver.click_ele(self.diagn_win, self.save_buttons[0], fail_save_txt)
        except Exception as click_save_e:
            self.failed_to_enter_code.append(code)
            return False
        else:
            sleep(.5)
            # If an alert appears, this dismisses it
            self.dismiss_alert()
            return

    # --------------------- Window Management ---------------------
    # Returns number of windows
    def num_windows(self):
        # Attempts to catch a failure to count windows. Could happen if too many attempts to count have been made quickly.
        try: return len(self.webdriver.driver.window_handles)
        except:
            sleep(5) # Waits and tries again (usually this will correct problem, since it arises from too many counts too quickly)
            try: return len(self.webdriver.driver.window_handles)
            # Still can't get window count even after waiting. Logs error and terminates import process.
            # Method that called this informs user and if relevant, logs all future codes as failed.
            except Exception as count_windows_e:
                self.ErrorLogger.log_error("Attempted to count windows with len(self.webdriver.window_handles) in num_windows()", count_windows_e)
                return False

    # Gets the most recently opened window (called when the new diagnosis window is needed)
    def get_diag_win(self):
        self.diagn_win = self.webdriver.driver.window_handles[-1]
        self.webdriver.switch_window(self.webdriver.main_win_handle, self.diagn_win)

    # Returns the title of the active window
    def active_win_title(self): return self.webdriver.driver.title

    # Closes whatever is the most recently opened window
    def close_most_rec_win(self):
        try:
            if self.webdriver.main_win_handle == self.webdriver.driver.window_handles[-1] or self.num_windows() == 1: return # Main window should not be closed
        except Exception as identify_window_handles_e: # An error could mean all windows are closed or something else
            self.ErrorLogger.log_error("Attempted to check main_win_handle against window_handles[-1] and attempted to count windows with num_windows()", identify_window_handles_e)
            return False
        try:
            self.webdriver.driver.switch_to.window(self.webdriver.driver.window_handles[-1])
            self.webdriver.driver.close()
        except: return False

    # Reduces windows to one (keeps self.webdriver.main_win_handle)
    def reduce_win_to_one(self):
        num_win = self.num_windows()

        if num_win == False: return False
        if num_win > 1:
            while self.num_windows() > 1:
                try: self.close_most_rec_win()
                except Exception as close_most_rec_e:
                    self.ErrorLogger.log_error("Attempted to close most recent window with close_most_rec_win()", close_most_rec_e)
                    break

    # Attempts to dismiss alert. No need to log error; try will fail if there was no alert in the first place.
    def dismiss_alert(self):
        sleep(.25)
        try: self.webdriver.driver.switch_to.alert.accept()
        except: pass

    # --------------------- Miscellaneous ---------------------
    def check(self, webd_obj: selenium.webdriver.remote.webelement.WebElement):
        if "Chateau" not in webd_obj.text:
            messagebox.showerror("Unlicensed facility", self.fail_to_verify)
            return False

    # Simple method keeping records of failing to enter data consistent
    def fail_other_log(self, code: str, element: str, field: str):
        self.failed_to_enter_other.append(f"{code} -- Failed to enter {element} in {field}")
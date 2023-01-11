import sys

if __name__ == "__main__":
    print("This is not the main module. Do not execute directly.")
    sys.exit()

import os.path, datetime
from pandas import read_csv, concat
from pandas import DataFrame
from tkinter import messagebox

class ErrorLogger:
    def __init__(self, log_file_name: str):
        self.headers = ["datetime", "notes", "exception"] # CSV headers
        self.log_file_name = log_file_name

    # Parameters are custom notes to be written to the log file and the exception thrown
    def log_error(self, notes: str, exception):
        file_exists = os.path.exists(self.log_file_name)

        if exception == None: return # Error that only needed to be reported to the user, not logged.

        # If user has the error logging file open, loop tells user they must close it.
        # Loop will never end until the file is closed.
        file_available = False
        while file_available == False:
            if file_exists == False:
                file_available = True
                continue

            try:
                testing = open(self.log_file_name, "a")
                testing.close()
            except PermissionError: # File is open
                messagebox.showerror(
                    "CM_error_log.csv is open",
                    "You must close the CM_error_log.csv file to continue."
                )
            # If the file cannot be read but it's not due to a permission error, the error handler cannot accommodate this.
            # No error logged in this case.
            except Exception as write_log_e:
                messagebox.showerror(
                    "Unknown error",
                    f"Could not access CM_error_log.csv. Note the below error and report it to the developer:\n\n{write_log_e}"
                )
                return
            else: file_available = True

        # Writes data to error logging file
        # Creates dataframe of new error to log
        df_new_err = DataFrame.from_dict({
            self.headers[0]: [datetime.datetime.now().strftime("%Y-%m-%d %H:%M")],
            self.headers[1]: [notes],
            self.headers[2]: [exception]
            })

        # If file does exist, concatenates existing data to new error to be logged
        # Otherwise will only log the original data
        if file_exists != False:
            try: df_err_log = read_csv(self.log_file_name, index_col = 0)
            except Exception as read_error_log_e: # Perhaps a corrupt file.
                messagebox.showerror(
                    "Failed to read error log",
                    "The error log could not be read. It may be corrupt.\n\n"
                        "The error log will be overwritten. Please manually save a backup of the current log before clicking OK."
                )
                try: os.remove(self.log_file_name)
                except: pass
                df_err_log = df_new_err
            else: df_err_log = concat([df_err_log, df_new_err])
        else: df_err_log = df_new_err

        # Overwrites original file (with either concatenated original data or with new log if new file created)
        try: df_err_log.to_csv(self.log_file_name)
        except Exception as fail_to_write_log_e:
            messagebox.showerror("Could not write to error log", "Could not write to error log. Please contact developer about this error.")

    """
    Deleting portions of the file if the file is > 50 error logs
    Keeps erorr logging file from becoming bloated over a high number of uses
    Only called when CodeM UP is closed out
    """
    def manage_file_size(self):
        file_exists = os.path.exists(self.log_file_name)
        if file_exists == False: return

        try:
            df_err_log = read_csv(self.log_file_name, index_col = 0)
            if len(df_err_log) > 50:
                df_err_log = df_err_log.iloc[len(df_err_log) - 50:]
                df_err_log.to_csv(self.log_file_name)
        except Exception as manage_file_size_e:
            messagebox.showerror("Error log failure", "Failed to manage error log file size. Please contact developer about this error.")
# CodeM UP
# aka, Codey McCodeface
# Software created by Matt Skaggs Ⓒ 2022
# This software is exclusively licensed for use at the Estates at Chateau in Minneapolis, MN

import sys

if __name__ != "__main__":
    print("This is the main module. Do not call.")
    sys.exit()

import csv
from tkinter_handler import CCWindow

# Reading in all settings to settings_dict
with open("CM_Settings.csv", "r") as cm_settings:
    reader = csv.reader(cm_settings)
    settings_list = [row for row in reader]
    settings_dict = {ele[0]: ele[1] for ele in settings_list}

# Instantiating main window (starting Codey McCodeface)
CC_main = CCWindow(
    settings_dict["pdf_default_dir"], # Directory user starts in when dialog window opens to import PDF
    settings_dict["excel_file_path"], # Excel file used to save and read code data
    settings_dict["excel_file_sheet_name"], # Name of the Excel sheet data is written to/read from
    int(settings_dict["excel_data_first_row"]), # The first row with data in it. Used to write data from PDF in pandas_handler.py.
    settings_dict["pcc_url"], # Page Chrome navigates to on launch
    int(settings_dict["window_x"]), # Main window width
    int(settings_dict["window_y"]), # Main window height
    settings_dict["template_file_path"], # Original template. Copied over CM_Codes.xlsx every time a PDF is imported.
    settings_dict["new_diag_button_x"], # New Diagnosis button xpath
    settings_dict["code_field_x"], # Field to enter ICD-10 code xpath
    settings_dict["code_desc_x"], # Field with PCC-generated code description xpath
    settings_dict["admis_date_id"], # ID for admission date in diag_win
    settings_dict["rank_x"], # Drop-down for Rank xpath
    settings_dict["clasif_x"], # Drop-down for Classification xpath
    settings_dict["comm_x"], # Comments field xpath
    settings_dict["confid_x"], # Confidential checkbox xpath
    settings_dict["log_file_name"], # Name of file to log errors to
    settings_dict["FID"]
)
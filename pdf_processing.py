import sys

if __name__ == "__main__":
    print("This is not the main module. Do not execute directly.")
    sys.exit()

import re
from os import path
from PyPDF2 import PdfFileReader
from full_codes import codes_list

# Imports PDFs and extracts ICD-10 codes from them
class PdfProcessing:
    def __init__(self):
        self.pdf_reader = None # Will be used for PyPDF2.PdfFileReader()
        self.pdf_all_text = [] # List of all text extracted from PDF (page text = liste ele)
        self.regex_results = [] # List of regex results from the entire PDF document

    # Returns a three-item tuple
    # tup[0] = either a list of all pages' text OR False if an error occurred
    # tup[1] = either None (successful text extraction) OR a string message to user about error
    # tup[2] = either None (successful text extraction or invalid PDF or no text in PDF) OR exception captured by try/except
    def import_pdf(self, pdf_path: str):
        self.pdf_all_text = [] # List of lists of codes. Each sub-list is from one page.

        # Checking for any unlikely problems reading the file
        # This catches invalid PDFs and should catch removed files
        try: self.pdf_reader = PdfFileReader(pdf_path)
        except Exception as read_pdf_e:
            return (False,
                    "PDF file could not be read.\n\nIs this a valid PDF, or was the file removed?",
                    read_pdf_e
                    )

        num_pages = self.pdf_reader.getNumPages()

        # Zero pages found in file. May be corrupt or not a PDF
        if num_pages == 0: return (False, "No pages were found in this file.\n\nIs this a valid PDF file?", None)

        # Extracts all text from all pages
        for page in range(0, self.pdf_reader.getNumPages()):
            self.pdf_all_text.append(self.pdf_reader.getPage(page).extract_text())

        # Checks if all pages have no text
        for page_text in self.pdf_all_text:
            # As soon as one element (page's extracted text) contains text, returns the results
            if page_text != "": return (self.pdf_all_text, None, None)

        # If method hasn't returned yet, then no text was found in file
        return (False,
                "No text found in document.\n\nDoes this PDF have text in it? Is it a scan or a fax?",
                None
                )

    # Receives list of text (each element = text from PDF page)
    # I referenced https://library.ahima.org/doc?oid=106177#.Y1XHwnbMK5c to build the regular expression to match the ICD-10 format
    def apply_regex(self, text_list: list):
        self.regex_results = [] # List of regex results from the entire PDF document

        # Checks for regex results in each list element (extracted page text)
        for page_text in text_list:
            page_results = re.findall("[A-Z]\d{2}\.?\w{0,4}", page_text)

            # Passes results to self.process_regex_page_results()
            self.process_regex_page_results(page_results)

        return sorted(self.regex_results)

    # Processes results from re.findall()
    # If any results, process_regex_page_results, checks them (actual codes, not already in list), and appends them to list self.regex_results
    def process_regex_page_results(self, page_results: list):
        if page_results == []: return
        for code in page_results:
            if code not in self.regex_results: # Checks this first because it's a quick way to rule out an already found code
                if code.replace(".", "") in codes_list: self.regex_results.append(code) # Is in master code list

    # Checks if Excel file is open (Excel creates hidden file prepended with "~$" when file is open)
    def is_excel_file_open(self, excel_path):
        if path.isfile("~$" + excel_path): return True
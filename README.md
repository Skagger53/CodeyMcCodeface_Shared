# Codey McCodeface (aka CodeM UP)
**Github files for CMN Splitter**

Auto enter ICD-10 codes into PCC

This software enters ICD-10 codes from an Excel spreadsheet into a patient profile in PointClickCare (PCC) using a webdriver.

To set up the codes to enter, the user may import codes automatically from a text-based PDF that contains the codes anywhere in its text. (All ICD-10 code matches will be imoprted.) This software will then enter these codes in the embedded Excel file. Alternatively the user may enter these codes into the Excel file by hand. Either way, the user may alter the list of codes as needed and add additional information in the Excel file table; the table columns correspond to the available options for codes in PCC.

The user does not have to ever open the Excel file; the user could import codes automatically and simply move on to entering them into PCC.

After the list of codes is prepared, the user will click a button to open a new intance of Chrome, where they can log into PCC and navigate to the relevant profile. Once the browser is at the profile page, the user can begin the upload process, and all codes will be entered into PCC. If any errors are encountered, these will be logged, and the software will inform the user of all errors after all codes have been entered.


**Libraries used**

*Tkinter (packaged with Python)*

Python Tkinter docs site:

https://docs.python.org/3/library/tkinter.html

Python history and license:

https://docs.python.org/3/license.html


*Selenium*
Main website

https://www.selenium.dev/

License

https://www.selenium.dev/documentation/about/copyright/#license


*Pyperclip*

Main site:

https://pypi.org/project/pyperclip/

Github:

https://github.com/asweigart/pyperclip

License:

https://github.com/asweigart/pyperclip/blob/master/LICENSE.txt


*PIL*

Stable release main website:

https://pillow.readthedocs.io/en/stable/

License:

https://pillow.readthedocs.io/en/stable/about.html#license


*Pandas*
Main website:

https://pandas.pydata.org/docs/

License:

https://github.com/pandas-dev/pandas/blob/main/LICENSE


*PyPDF2*

Main website:

https://pypi.org/project/PyPDF2/

License:

https://github.com/py-pdf/pypdf/blob/main/LICENSE


*Openpyxl*

Stable version main site:

https://openpyxl.readthedocs.io/en/stable/

License:

https://github.com/theorchard/openpyxl/blob/master/LICENCE.rst

# Spreadsheet-Unifier

The basic legal stuff!
    By running this software you agree to the attached License document


What is required to run this program:
    You must first ensure that you have Python installed on your system. This software will work with either the Anaconda bundle, or the normal Python download from python.org
    Please note if you are using the standard Python download you must also download the openpyxl module from pip

What the program does:
    This program takes multiple spreadsheets with the same or even different number of columns and merges them into one spreadsheet while also merging row data, removes duplicates, and strips out leading and trailing spaces.

limitations:
    This program currently can only work with spreadsheets with a single sheet in each file. If your spreadsheet (like an excel workbook) has multiple sheets, please break them out first.
    Currently this product can only handle Excel .xlsx / .xls files and CSV files. Please convert any other spreadsheets.

 
How to use:
    Ensure you have a Python Interpreter installed.
        If you are an individual, or if your organization has a license with Anaconda, you can download Python and all the related packages simply by downloading the Anaconda package from Anaconda.com
        If you are a business, Nonprofit, or individual who wishes to not use the Anaconda package you can install python for free and for any use case from python.org
	        If you download and install python from this method, you must also install the openpyxl module which is used to read and write excel files.
		        To do this once python is install you will open your terminal (on windows this can be done by either CMD or Powershell) and type pip install openpyxl
    Once python is installed and openpyxl you can run the software by opening spreadsheet unifier.py
    This software requires all the files you wish to merge into one spreadsheet be in a single folder.

# Spreadsheet-Unifier

The basic legal stuff!
    By running this software you agree to the attached License document


What is required to run this program:
    You must first ensure that you have Python installed on your system. this software will work with either the Anaconda bundle, or the normal Python download
    Please note if you are using the standard Python download you must also download an additional module from pip
    This module is openpyxl

What the program does:
    This program takes multiple spreadsheets with the same or even different number of columns and merges them into one spreadsheet while also merging row data and removing duplicates. This tool will also strip out leading and trailing spaces.

limitations:
    This program currently can only work with spreadsheets with a single sheet
    Currenlty this product can only handle Excel files (.xlsx) and CSV files. Please convert any other spreadsheets.

 
How to use:
    Ensure you have a Python Interpreter installed.
        If you are an individual you can download python and all the related packages simply by downloading the Anaconda package from Anaconda.com
        If you are a business, Nonprofit, or individual who wishes to not use the Anaconda package you can install python for free and for any use case from python.org
	        If you download and install python from this method you must also install the openpyxl module which is used to read and write excel files.
		        To do this once python is install you will open your terminal ( on windows this can be done by either CMD or Powershell) and type pip install openpyxl
    Once python is installed and openpyxl you can run the software by opening spreadsheet unifier.py
    This software requires all the files you wish to merge in to one speadsheet be in a single folder.

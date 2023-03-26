# Spreadsheet-Unifier

The basic legal stuff!
    By running this software you agree to the attached License document


  Disclaimer of Warranty.

  THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY
APPLICABLE LAW.  EXCEPT WHEN OTHERWISE STATED IN WRITING THE COPYRIGHT
HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM "AS IS" WITHOUT WARRANTY
OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO,
THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
PURPOSE.  THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM
IS WITH YOU.  SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF
ALL NECESSARY SERVICING, REPAIR OR CORRECTION.

  Limitation of Liability.

  IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING
WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MODIFIES AND/OR CONVEYS
THE PROGRAM AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES, INCLUDING ANY
GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE
USE OR INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED TO LOSS OF
DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY YOU OR THIRD
PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER PROGRAMS),
EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE POSSIBILITY OF
SUCH DAMAGES.



What is required to run this program.
    You must first ensure that you have Python installed on your system. this software will work with either the Anaconda bundle, or the normal Python download
        • Please note if you are using the standard Python download you must also download an additional module from pip
            o This module is openpyxl

What the Program does
    This program takes multiple spreadsheets with the same or even different number of columns and merges them into one spreadsheet while also merging row data and removing duplicates. This tool will also strip out leading and trailing spaces.

limitations
    This program currently can only work with spreadsheets with a single sheet
    Currenlty this product can only handle Excel files (.xlsx) and CSV files. Please convert any other spreadsheets.

 
How to use
    Ensure you have a Python Interpreter installed.
        If you are an individual you can download python and all the related packages simply by downloading the anaconda package from Anaconda.com

        If you are a business, Nonprofit, or individual who wishes to not use the anaconda package you can install python for free and for any use case from python.org
	        If you download and install python from this method you must also install the openpyxl module which is used to read and write excel files.
		        To do this once python is install you will open your terminal ( on windows this can be done by either CMD or Powershell) and type pip - -install openpyxl

    Once python is installed and openpyxl you can run the software by opening spreadsheet unifier.py
    This software requires all the files you wish to merge in to one speadsheet be in a single folder.

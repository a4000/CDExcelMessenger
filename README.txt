CDExcelMessenger.py has functions that allow passing data between an Excel file and a Compound Discoverer (CD) results file. 
CDExcelNotebook.ipynb is a Jupyter Notebook that is designed to make it easy to use the functions in CDExcelMessenger.py
This project is still not complete. More testing needs to be done and additional functionality is planned.

The code in this project has been written by Adam Bennett

The main functions in CDExcelMessenger.py are updateCDResultsFile(), and updateExcelFile(). 
These are the functions that get called in CDExcelNotebook.ipynb to pass data between the Excel file and the CD results file.
These functions call other functions in CDExcelMessenger.py.

updateCDResultsFile() lets the user import data from an Excel file into a CD results file. 
The user can add new columns to CD or update certain columns already in CD (Tags, Checked, Name, and any columns previously added by CDExcelMessenger.py). 
The Tags column in CD is 15 boxes that can be used as flags by the user. 
The updateCDResultsFile() function lets the user automate the flagging of these Tag boxes using columns in The Excel file. 
This function also lets the user choose the threshold values that get used when flagging the Tag boxes, 
and the user can choose the number of Tag boxes that are visable when opening the CD results file in the CD software.

updateExcelFile() lets the user import data from a CD results file into an Excel file.
This function doesn't allow getting certain columns from CD (e.g. the Area columns).
The user can still use this function on any column that is compatible with the updateCDResultsFile() function.
This allows the user to interact with an Excel file and CD results file in an iterative workflow.

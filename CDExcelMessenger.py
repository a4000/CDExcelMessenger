#####################################################################################
## Import Modules
#####################################################################################

import sqlite3
import os.path
import re
import logging
logger = logging.Logger('catch_all')

try:
    import pandas as pd
except:
    print("\nError: Pandas can't be imported")


#####################################################################################
## Function: formatStringToSQLite()
#####################################################################################  
'''
This function converts a string into a string that can be used as an SQLite column name
'string' = The string that needs to be converted to an SQLite friendly format
''' 
    
def formatStringToSQLite(string):

    # Strip characters that aren't alphanumeric or underscores
    string = re.sub(r'\W+', '', string)
        
    # Strip numbers at the start of the column
    string = string.lstrip("0123456789") 
    
    return string
    
    
#####################################################################################
## Function: getTagBytes()
#####################################################################################
'''
This function gets the string stored in the Tags column of an Excel file and converts that 
string into bytes to be stored in the CD results file
'tagString' = the string that needs to be converted to bytes
'''

def getTagBytes(tagString):
    try:
        # If there are any Tags
        if not pd.isna(tagString):
            # Remove whitespace
            tagString = tagString.replace(" ", "")
            # Create a list of tags
            tagList = tagString.split(";")
            
            # Make sure each tag is in uppercase
            for i in range(len(tagList)):
                tagList[i] = tagList[i].upper()
            # Put Tags in a Set to remove duplicates
            tagSet = set(tagList)

            # Create the tag byte string in the correct format for the results file
            tagBytes = b""
            if "A" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "B" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "C" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "D" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "E" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "F" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "G" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "H" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "I" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "J" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "K" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "L" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "M" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "N" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if "O" in tagSet:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            
            return tagBytes
        
        # There is no tags
        else:
            return -1
    
    # Get info about other errors
    except Exception as e:
        logger.error(e, exc_info=True)
        return -1
        
        
#####################################################################################
## Function: convertTagBytesToString()
#####################################################################################
'''
This function gets the bytes stored in the Tags column of a CD results file and converts those 
bytes into a string to be stored in the Excel file
'tagBytes' = the bytes that needs to be converted to a string
'''

def convertTagBytesToString(tagBytes):
    tagString = ""
    tagBytes = str(tagBytes)
    tagBytes = tagBytes[2:-1]
    
    if tagBytes[:8] == "\\x01\\x01":
        tagString = tagString + "A;"
    if tagBytes[8:16] == "\\x01\\x01":
        tagString = tagString + "B;"
    if tagBytes[16:24] == "\\x01\\x01":
        tagString = tagString + "C;"
    if tagBytes[24:32] == "\\x01\\x01":
        tagString = tagString + "D;"
    if tagBytes[32:40] == "\\x01\\x01":
        tagString = tagString + "E;"
    if tagBytes[40:48] == "\\x01\\x01":
        tagString = tagString + "F;"
    if tagBytes[48:56] == "\\x01\\x01":
        tagString = tagString + "G;"
    if tagBytes[56:64] == "\\x01\\x01":
        tagString = tagString + "H;"
    if tagBytes[64:72] == "\\x01\\x01":
        tagString = tagString + "I;"
    if tagBytes[72:80] == "\\x01\\x01":
        tagString = tagString + "J;"
    if tagBytes[80:88] == "\\x01\\x01":
        tagString = tagString + "K;"
    if tagBytes[88:96] == "\\x01\\x01":
        tagString = tagString + "L;"
    if tagBytes[96:104] == "\\x01\\x01":
        tagString = tagString + "M;"
    if tagBytes[104:112] == "\\x01\\x01":
        tagString = tagString + "N;"
    if tagBytes[112:120] == "\\x01\\x01":
        tagString = tagString + "O;"
    # Remove the last ;
    if len(tagString) > 0:
        tagString = tagString[:-1]
        
    return tagString
        
    
#####################################################################################
## Function: getDataFromExcelFile()
#####################################################################################
'''
This function gets the data from an Excel file. 
If a column is not the Tags column, then empty rows are filled with default values
'excelFilePath' = The path to an Excel file
'''

def getDataFromExcelFile(excelFilePath):
    try:
        excelData = pd.read_excel(excelFilePath)
                
        for column in excelData:
                    
            if column != "Tags":
                # If the column has string values
                if excelData.dtypes[column] == "object":
                    excelData[column] = excelData[column].fillna("")
                # If the column has boolean values
                elif excelData.dtypes[column] == "bool":
                    excelData[column] = excelData[column].fillna(False)
                # If the column has int, or float values
                else:
                    excelData[column] = excelData[column].fillna(0)
            
        return excelData
            
    # Value Error can be caused by the Excel file not having the correct column names
    except ValueError:
        print("\nValue Error")
        print("Make sure "+excelFilePath+" has these columns "+colNameList)
        # Return an empty dataframe
        return pd.DataFrame() 
        
    # If the Excel file can't be found
    except FileNotFoundError:
        print("\nFileNotFoundError")
        print("Can't find "+excelFilePath)
        # Return an empty dataframe
        return pd.DataFrame()
    
    # If permission to the Excel file was denied
    except PermissionError:
        print("\nPermissionError")
        print("Couldn't gain permission to the Excel File")
        print("Make sure "+excelFilePath+" is not open in another program")
        # Return an empty dataframe
        return pd.DataFrame()
    
    # Get info about other errors
    except Exception as e:
        logger.error(e, exc_info=True)
        # Return an empty dataframe
        return pd.DataFrame()







#####################################################################################
## Function: updateDataInCDResultsFile()
#####################################################################################
'''
    This function imports data from an Excel file to a Compound Discoverer (CD) results file.
    'cdResultsFilePath' = The path to a CD results file
    'excelFilePath' = The path to an Excel file
    'updateColNameList' = a list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
        all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file)
    'hideMessages' = Boolean value that controls whether or not the print statements get run. 
        Set to True to hide messages and return a list of outputs (default is False)
'''
    
def updateDataInCDResultsFile(cdResultsFilePath, excelFilePath, updateColNameList = None, hideMessages = False):

    cdResultsFileName = cdResultsFilePath.rpartition("/")[2]
    excelFileName = excelFilePath.rpartition("/")[2]
    
    if hideMessages == False: 
        print("\nImporting data from "+excelFileName+" into "+cdResultsFileName)
    else:
        outputList = ["Importing data from "+excelFileName+" into "+cdResultsFileName]
    try:
        # Get data from the Excel File and store as a dataframe
        excelData = getDataFromExcelFile(excelFilePath)
        
        # If there was no exceptions raised during the getDataFromExcelFile() function
        if not excelData.empty:
        
            # If the results file can be found
            if os.path.exists(cdResultsFilePath):
            
                # Open connection to the Compound Discoverer File
                conn = sqlite3.connect(cdResultsFilePath)
                cursor = conn.cursor()

                # Get number of rows in the Excel file
                excelRowCount = len(excelData.index)
            
                # colNameTupleList is going to be a list of tuples. The tuples are going to contain the column name in the SQLite format and the Display name
                colNameTupleList = []
            
                # Get the ID of compound table
                cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
                compoundTblID = cursor.fetchall()[0][0] 

                # If the user has provided a list of columns to update
                if updateColNameList is not None:
                    for colDisplayName in updateColNameList:
                        if colDisplayName in excelData.columns:
                            # colSQLiteName contains the display name that's been converted to an SQLite friendly format
                            colSQLiteName = formatStringToSQLite(colDisplayName)
                        
                            # This is the tuple that gets added to colNameTupleList
                            colNameTuple = (colSQLiteName, colDisplayName)
                    
                            # Check if 'ConsolidatedUnknownCompoundItems' already has the column
                            cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='"+colSQLiteName+"';")
                        
                            # If the column doesn't exist in the CD results file
                            if cursor.fetchall()[0][0] == 0:
                                # Add the colNameTuple to the colNameTupleList which will be used later when updating data in the CD Results file
                                colNameTupleList.append(colNameTuple)
                            
                                # Get the data type of the column
                                colDataType = excelData.dtypes[colDisplayName]
                           
                                # Add new column and set default values based on the data type of the columns
                                # Also set customDataType and valueType which gets used in the CD database
                                if colDataType == "int64":
                                    # Add column to compound table
                                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colSQLiteName+" INTEGER;")
                                    # Add default values to the new column
                                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colSQLiteName+" = (?);", (str(0), ))
                                
                                    customDataType = "2"
                                    valueType = "A170C73A-BD79-493B-B24A-B981BAF6DCC5"
                                    #IF THIS DOESN'T WORK, TRY customDataType 1 and valueType B186F5FB-FD41-4087-8A54-CB52CCA0E3DF
                                
                                elif colDataType == "float64":
                                    # Add column to compound table
                                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colSQLiteName+" REAL;")
                                    # Add default values to the new column
                                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colSQLiteName+" = (?);", (str(0.0), ))
                                
                                    customDataType = "3"
                                    valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                                elif colDataType == "bool":
                                    # Add column to compound table
                                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colSQLiteName+" NUMERIC;")
                                    # Add default values to the new column
                                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colSQLiteName+" = (?);", ("False", ))
                                
                                    customDataType = "4"
                                    valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                                elif colDataType == "object":
                                    # Add column to compound table
                                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colSQLiteName+" TEXT;")
                                    # Add default values to the new column
                                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colSQLiteName+" = (?);", ("", ))
                                
                                    customDataType = "4"
                                    valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                                else:
                                    # Add column to compound table
                                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colSQLiteName+" TEXT;")
                                    # Add default values to the new column
                                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colSQLiteName+" = (?);", ("", ))
                                
                                    customDataType = "4"
                                    valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                            
                                #[CustomDataType, and ValueType will need to set depending on data type]
                                # Add column details to columns table
                                cursor.execute("INSERT INTO DataTypesColumns \
                                                (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType, \
                                                Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                                                Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                                                Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                                                Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                                                VALUES \
                                                ("+str(compoundTblID)+", (?), (?), 1, (?), \
                                                0, -1, '', (?), (?), \
                                                '', 1, '',\
                                                4, 0, -1, \
                                                '', 1, 0);", (colSQLiteName, customDataType, valueType, colDisplayName, colDisplayName+": This column has been added by CDExcelMessenger.py"))
                                if hideMessages == False: 
                                    print("Column: "+colDisplayName+" added to "+cdResultsFileName)
                                else:
                                    outputList.append("Column: "+colDisplayName+" added to "+cdResultsFileName)
                        
                            # If the column already exists in the CD results file
                            else:
                                cursor.execute("SELECT Grid_AllowEdit FROM DataTypesColumns WHERE DataTypeID = "+str(compoundTblID)+" AND DBColumnName = (?);", (colSQLiteName, ))
                                # If the column is editable
                                if cursor.fetchall()[0][0] == 1:
                            
                                    #[May need to check if data type has changed]
                            
                                    colNameTupleList.append(colNameTuple)
                                # If the column is not editable
                                else:
                                    if hideMessages == False: 
                                        print("WARNING: "+colDisplayName+" is already used in "+cdResultsFileName+" and it is a non-editable column, give this column a different name if you want to add this column")
                        else:
                            if hideMessages == False: 
                                print("WARNING: "+colDisplayName+" can't be found in "+excelFileName)
                            else:
                                outputList.append("WARNING: "+colDisplayName+" can't be found in "+excelFileName)
                # User wants to update all editable columns 
                else: 
                    # Gets all editable columns (Tags, Checked, Name, and columns added by user)
                    # colNameTupleList is a list of tuples
                    cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = "+str(compoundTblID)+" AND Grid_AllowEdit = 1;")
                    colNameTupleList = cursor.fetchall()
            
                # Check if 'ConsolidatedUnknownCompoundItems' has the 'Cleaned' column
                cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='Cleaned';")
                # If the 'Cleaned' column doesn't exist, create it
                if cursor.fetchall()[0][0] == 0:
                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN Cleaned")
                    # Add Cleaned details to columns table
                    cursor.execute("INSERT INTO DataTypesColumns \
                                    (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType,\
                                    Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                                    Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                                    Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                                    Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                                    VALUES \
                                    ("+str(compoundTblID)+", 'Cleaned', 4, 1, '3245F562-3044-4BC0-9091-3813CA7AE5BC', \
                                    0, -1, '', 'Cleaned', 'Shows the rows that have been updated with CDExcelMessenger.py', \
                                    '', 1, '',\
                                    4, 0, -1, \
                                    '', 0, 0);")
                    # Set Cleaned to False for all rows
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET Cleaned = 'False';") 
                # The 'Cleaned' column does exist
                else:
                    # Check if 'ConsolidatedUnknownCompoundItems' has the 'OldCleaned' column
                    cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='OldCleaned';")
                    # If the 'OldCleaned' column doesn't exist
                    if cursor.fetchall()[0][0] == 0:
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems RENAME COLUMN Cleaned to OldCleaned")
                        # Add Old Cleaned details to columns table
                        cursor.execute("INSERT INTO DataTypesColumns \
                                        (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType, \
                                        Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                                        Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                                        Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                                        Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                                        VALUES \
                                        ("+str(compoundTblID)+", 'OldCleaned', 4, 1, '3245F562-3044-4BC0-9091-3813CA7AE5BC', \
                                        0, -1, '', 'Old Cleaned', 'Shows the rows that have been updated with CDExcelMessenger.py', \
                                        '', 1, '',\
                                        4, 0, -1, \
                                        '', 0, 0);")
                
                        # Add a new column called 'Cleaned'
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN Cleaned")
                
                        # Set Cleaned to False for all rows
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET Cleaned = 'False';") 
                
                    # The OldCleaned column does exist    
                    else:
                        # Copy current cleaned column into the old cleaned column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET OldCleaned = Cleaned")
                    
                        # Set Cleaned to False for all rows
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET Cleaned = 'False';") 
            
                # Loop through each tuple in the column names list
                for colNameTuple in colNameTupleList:
                    colSQLiteName = colNameTuple[0]
                    colDisplayName = colNameTuple[1]
                    if colDisplayName in excelData.columns:
                        # Loop through each row in the Excel file
                        for row in range(excelRowCount):
                            # The Tags column needs to be handled differently
                            if colSQLiteName == "Tags":
                                value = getTagBytes(excelData.at[row,colDisplayName])
                                if value == -1:
                                    value = b"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"
                            
                                # Molecular weight and retention time is needed to match rows between the excel file and CD results file
                                MW = excelData.at[row,"Calc. MW"]
                                RT = excelData.at[row,"RT [min]"]
                            
                                # Update the current row and column in the CD results file, also set Cleaned to True
                                cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colSQLiteName+" = (?), Cleaned = 'True' WHERE ROUND(MolecularWeight, 5) = ROUND("+str(MW)+", 5) and ROUND(RetentionTime, 3) = ROUND("+str(RT)+", 3);", (value, ))     
                            
                                # Raise an exception if multiple rows in the CD results file matched with a row in the Excel file
                                if cursor.rowcount > 1:
                                    if hideMessages == False:
                                        print("WARNING: multiple rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                                    else:
                                        outputList.append("WARNING: multiple rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                                    raise Exception
                                    
                            # If the column isn't the Tags column
                            else:
                                # The Checked column also needs to be handled differently
                                if colSQLiteName == "Checked": 
                                    value = excelData.at[row,colDisplayName]
                                    if bool(value):
                                        value = 1
                                    else:
                                        value = 0
                                # Not the Tags or Checked column
                                else:
                                    value = excelData.at[row,colDisplayName]
                            
                            
                                # Molecular weight and retention time is needed to match rows between the excel file and CD results file
                                MW = excelData.at[row,"Calc. MW"]
                                RT = excelData.at[row,"RT [min]"]
                            
                                # Update the current row and column in the CD results file, also set Cleaned to True
                                cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colSQLiteName+" = (?), Cleaned = 'True' WHERE ROUND(MolecularWeight, 5) = ROUND("+str(MW)+", 5) and ROUND(RetentionTime, 3) = ROUND("+str(RT)+", 3);", (str(value), ))     
                            
                                # Raise an exception if multiple rows in the CD results file matched with a row in the Excel file
                                if cursor.rowcount > 1:
                                    if hideMessages == False:
                                        print("WARNING: multiple rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                                    else:
                                        outputList.append("WARNING: multiple rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                                    raise Exception
                        
                        if hideMessages == False: 
                            print("Column: "+colDisplayName+" found in "+excelFileName)
                        else:
                            outputList.append("Column: "+colDisplayName+" found in "+excelFileName)
                    else:
                        if hideMessages == False: 
                            print("WARNING: "+colDisplayName+" can't be found in "+excelFileName)
                        else:
                            outputList.append("WARNING: "+colDisplayName+" can't be found in "+excelFileName)
                
                conn.commit()
                if hideMessages == False: 
                    print(cdResultsFileName+" updated")
                else:
                    outputList.append(cdResultsFileName+" updated")
                # Close the connection to the Compound Discoverer file
                cursor.close()
                conn.close()
    
            # If the results file can't be found
            else:
                if hideMessages == False: 
                    print("WARNING: "+cdResultsFilePath+" can't be found")
                else:
                    outputList.append("WARNING: "+cdResultsFilePath+" can't be found")
        
        if hideMessages == True:
            return outputList      
    
    # Get info about errors
    except Exception as e:
        if hideMessages == False: 
            logger.error(e, exc_info=True)
        
        
#####################################################################################
## Function: updateDataInExcelFile()
#####################################################################################
'''
    This function imports data from a Compound Discoverer (CD) results file into an Excel file.
    'cdResultsFilePath' = The path to a CD results file
    'excelFilePath' = The path to an Excel file
    'updateColNameList' = a list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
        all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file)
    'hideMessages' = Boolean value that controls whether or not the print statements get run. 
        Set to True to hide messages and return a list of outputs (default is False)
'''
    
def updateDataInExcelFile(cdResultsFilePath, excelFilePath, updateColNameList = None, hideMessages = False):
    
    cdResultsFileName = cdResultsFilePath.rpartition("/")[2]
    excelFileName = excelFilePath.rpartition("/")[2]
    
    if hideMessages == False: 
        print("\nImporting data from "+cdResultsFileName+" into "+excelFileName)
    else:
        outputList = ["Importing data from "+cdResultsFileName+" into "+excelFileName]
    try: 
        excelData = getDataFromExcelFile(excelFilePath)

        # If there was no exceptions raised during the getDataFromExcelFile() function
        if not excelData.empty:

            # If the results file can be found
            if os.path.exists(cdResultsFilePath):
            
                # Open connection to the Compound Discoverer File
                conn = sqlite3.connect(cdResultsFilePath)
                cursor = conn.cursor()

                # Get number rows in the Excel file
                excelRowCount = len(excelData.index)
            
                # colNameTupleList is going to be a list of tuples. The tuples are going to contain the column name in the SQLite format and the Display name
                colNameTupleList = []
            
                # Get ID of compound table
                cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
                compoundTblID = cursor.fetchall()[0][0] 
            
                # If the user has provided a list of columns to update
                if updateColNameList is not None:
                    for colDisplayName in updateColNameList:
                    
                        # colSQLiteName contains the display name that's been converted to an SQLite friendly format
                        colSQLiteName = formatStringToSQLite(colDisplayName)
                    
                        # Check if 'ConsolidatedUnknownCompoundItems' has the column
                        cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name=(?);", (colSQLiteName, ))
                        # If the column exists
                        if cursor.fetchall()[0][0] == 1:
                        
                            # This is the tuple that gets added to colNameTupleList
                            colNameTuple = (colSQLiteName, colDisplayName)
                        
                            # Add the colNameTuple to the colNameTupleList which will be used later when updating data in the Excel file
                            colNameTupleList.append(colNameTuple)
                        
                        # If the column doesn't exist
                        else:
                            if hideMessages == False: 
                                print("WARNING: "+colDisplayName+" can't be found in "+cdResultsFileName)
                            else:
                                outputList.append("WARNING: "+colDisplayName+" can't be found in "+cdResultsFileName)
            
                # User wants to update all editable columns 
                else:
                    # Gets all editable columns (Tags, Checked, Name, and columns added by user)
                    # colNameTupleList will be a list of tuples
                    cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = "+str(compoundTblID)+" AND Grid_AllowEdit = 1;")
                    colNameTupleList = cursor.fetchall()
        
                # Loop through each tuple in the column names list
                for colNameTuple in colNameTupleList:
                    colSQLiteName = colNameTuple[0]
                    colDisplayName = colNameTuple[1]
                
                    # If the column is not in the excel file, add the column with placeholder values
                    if colDisplayName not in excelData.columns:
                        excelData[colDisplayName] = [0]*excelRowCount
                        if hideMessages == False: 
                            print("Column: "+colDisplayName+" added to "+excelFileName)
                        else:
                            outputList.append("Column: "+colDisplayName+" added to "+excelFileName)
                    else:
                        if hideMessages == False: 
                            print("Column: "+colDisplayName+" found in "+excelFileName)
                        else:
                            outputList.append("Column: "+colDisplayName+" found in "+excelFileName)
                            
                    # Check if the column data type is boolean, the colIsBool variable gets used to make sure the column stays as bool
                    colIsBool = False
                    if excelData.dtypes[colDisplayName] == "bool":
                        colIsBool = True
                
                    # Loop through each row in the excel file
                    for row in range(excelRowCount):
                    
                        # Molecular weight and retention time is needed to match rows between the excel file and CD results file
                        MW = excelData.at[row,"Calc. MW"]
                        RT = excelData.at[row,"RT [min]"]
                        
                        # Get the value of the current row and column from the CD results file    
                        cursor.execute("SELECT "+colSQLiteName+" FROM ConsolidatedUnknownCompoundItems WHERE ROUND(MolecularWeight, 5) = ROUND("+str(MW)+", 5) and ROUND(RetentionTime, 3) = ROUND("+str(RT)+", 3);")     
                        selectStatementResults = cursor.fetchall()
                        value = selectStatementResults[0][0]
                    
                        # Raise an exception if multiple rows in the CD results file matched with a row in the Excel file
                        if len(selectStatementResults) > 1:
                            if hideMessages == False: 
                                print("WARNING: multiple rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                            
                        # If exactly one row in the CD results file matched with a row in the Excel file
                        elif len(selectStatementResults) == 1:
                            # The Tags column needs to be handled differently
                            if colSQLiteName == "Tags":
                                # Update the current row and column in the Excel data frame after converting the bytes value to a string
                                excelData.at[row,colDisplayName]=convertTagBytesToString(value)
    
                            # If the value is not stored as bytes
                            elif type(value) != bytes:
                                # if the current column is a boolean column
                                if colIsBool:
                                    # If the value is stored as a string
                                    if type(value) == str:
                                        if value.upper() == "TRUE":
                                            value = True
                                        else:
                                            value = False
                                    # If the value is not stored as a string
                                    else:
                                        value = bool(value)
                                # Update the current row and column in the Excel data frame
                                excelData.at[row,colDisplayName]=value
                            
                            # If the current column stores bytes values, but is not the Tags column
                            else:
                                if hideMessages == False: 
                                    print("WARNING: "+colDisplayName+" can't be updated")
                                    break
                                else:
                                    outputList.append("WARNING: "+colDisplayName+" can't be updated")
                                    break
                        # No rows in the CD results file match with the current row in the Excel file
                        else:
                            if hideMessages == False: 
                                print("WARNING: no rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                            else:
                                outputList.append("WARNING: no rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                    # This is to make sure boolean values are set as bool in the Excel file
                    if colIsBool:                 
                        excelData[colDisplayName] = excelData[colDisplayName].astype('bool')
                # Update the Excel file with the Excel data frame
                excelData.to_excel(excelFilePath, index=False)  
                
                if hideMessages == False: 
                    print(excelFileName+" updated")
                else:
                    outputList.append(excelFileName+" updated")

                # Close the connection to the Compound Discoverer file
                cursor.close()
                conn.close()
           
           # If the results file can't be found
            else:
                if hideMessages == False: 
                    print("WARNING: "+cdResultsFilePath+" can't be found")
                else:
                    outputList.append("WARNING: "+cdResultsFilePath+" can't be found")
        
        if hideMessages == True:
            return outputList 
        
    # Get info about errors
    except Exception as e:
        if hideMessages == False: 
            logger.error(e, exc_info=True)
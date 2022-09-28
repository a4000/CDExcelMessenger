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
## Function: formatStringToSQLiteColumn()
#####################################################################################  
'''
This function converts a string into a string that can be used as an SQLite column name
'string' = The string that needs to be converted to an SQLite friendly format
''' 
    
def formatStringToSQLiteColumn(string):

    # Strip characters that aren't alphanumeric or underscores
    string = re.sub(r'\W+', '', string)
        
    # Strip numbers at the start of the column
    string = string.lstrip("0123456789") 
    
    return string
    
   
#####################################################################################
## Function: changeTagNamesInCD()
#####################################################################################
'''
This function changes the name of tags in the CD results file 
based on the names of boolean columns in the Excel file, 
if the user has selected that column in 'optionalTagList'

This function also returns a list of user chosen tags if the
Excel file has a column with that name and if that column is boolean

'cursor' = an SQLite cursor
'excelData' = a DataFrame containing Excel data
'excelFilePath' = the path to an Excel file
'optionalTagList' = a list of boolean columns in the Excel file that the user wishes to set as tags in CD), 
    this function uses the Excel column names to change the name of tags in the CD results file
'hideMessages' = Boolean value that controls whether or not the print statements get run. 
    Set to True to hide messages and return a list of outputs (default is False)
'''

def changeTagNamesInCD(cursor, excelData, excelFilePath, optionalTagList, hideMessages):
    try:
        # 'ID' is used to match with the tags in the CD results file 
        ID = 1
        
        # This variable is the output of this function
        # It stores the strings in 'optionalTagList' that are boolean columns in the Excel file
        tagListOutput = []
        
        # Loop through each string that the user wants to use as tag names 
        for optionalTag in optionalTagList:
            # if the tag name is a column in the Excel file
            if optionalTag in excelData.columns:
                # if the column in the Excel file is boolean
                if excelData.dtypes[optionalTag] == "bool":
                    # change Name and Description in DataDistributionBoxes WHERE BoxID = ID
                    cursor.execute("UPDATE DataDistributionBoxes SET Name = (?), Description = (?) WHERE BoxID = "+str(ID)+";", (optionalTag, "Matching entry in: "+optionalTag+".")) 
                    
                    # Set the tags visability to True
                    cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'True' WHERE BoxID = "+str(ID)+";") 
                    ID = ID + 1
                    tagListOutput.append(optionalTag)
                    
                # if the column in the Excel file is not boolean
                else:
                    print("WARNING: "+optionalTag+" is not the Boolean type")
                    
            # if the tag name is not a column in the Excel file
            else:
                print("WARNING: "+optionalTag+" not found in "+excelFilePath)
        
        # Loop through the tag IDs for tags that didn't have their names changed
        while ID < 16:
            # Set the tags visability to False
            cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'False' WHERE BoxID = "+str(ID)+";") 
            ID = ID + 1
            
        if tagListOutput != {}:
            return tagListOutput
        else:
            print("WARNING: none of the strings in "+optionalTagList+" are boolean columns in "+excelFilePath)
            return None
    
    # Get info about errors
    except Exception as e:
        logger.error(e, exc_info=True)
        return -1
    
    
#####################################################################################
## Function: getTagBytes()
#####################################################################################
'''
This function gets the string stored in the Tags column of an Excel file and converts that 
string into bytes to be stored in the CD results file
'tagString' = the string that needs to be converted to bytes
'cdResultsFilePath' = the path to a CD results file
'cursor' = an SQLite cursor
'''

def getTagBytes(tagString, cdResultsFilePath, cursor):
    try:
        # If there are any Tags
        if not pd.isna(tagString):
            
            # Create a list of tags, ';' is the delimiter
            tagList = tagString.split(";")
            
            # Loop through the tag list to remove whitespace from the left and right of the tag
            for i in range(len(tagList)):
                tagList[i] = tagList[i].strip()
            
            # 'keys' will contain Tags found in the tagString
            # 'keys' will be used to match rows in the SQLite statement
            keys = ""
            # Loop through the Tags list
            for tag in tagList:
                # Add the Tags to keys
                keys += "'"+str(tag)+"', "
            
            # Remove last ', ' from the keys
            keys = keys[:-2]
            
            # get the ID of the Tags in 'keys' (the Tags that were found in the tagString) 
            cursor.execute("SELECT BoxID FROM DataDistributionBoxes WHERE Name IN ("+keys+");")
            tempIDList = cursor.fetchall()
            IDList = []
            for i in tempIDList:
                IDList.append(i[0])

            # 'tagBytes' will be the output of this function
            tagBytes = b""
            
            # Create the tag byte string in the correct format for the results file
            # if an ID is found in the IDList add \x01\x01 to indicate that the Tag with that ID is checked
            # \x00\x00 is used to idicate that the Tag with that ID is not checked
            if 1 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 2 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 3 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 4 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 5 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 6 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 7 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 8 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 9 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 10 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 11 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 12 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 13 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 14 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            if 15 in IDList:
                tagBytes += b"\x01\x01"
            else:
                tagBytes += b"\x00\x00"
            
            return tagBytes
        
        # There is no tags
        else:
            return -1
    
    # Get info about errors
    except Exception as e:
        logger.error(e, exc_info=True)
        return -1
        
        
#####################################################################################
## Function: getTagString()
#####################################################################################
'''
This function gets the bytes stored in the Tags column of a CD results file and converts those 
bytes into a string to be stored in the Excel file
'tagBytes' = the bytes that needs to be converted to a string
'cdResultsFilePath' = the path to a CD results file
'cursor' = an SQLite cursor
'''

def getTagString(tagBytes, cdResultsFilePath, cursor):
    try:    
        # Get a list of IDs for the Tags that are visable in CD
        cursor.execute("SELECT BoxID FROM DataDistributionBoxExtendedData WHERE ValueString = 'True';")
        tempBoxIDList = cursor.fetchall()
        boxIDList = []
        for boxID in tempBoxIDList:
            boxIDList.append(boxID[0])
    
        # get the names of the Tags
        cursor.execute("SELECT Name FROM DataDistributionBoxes WHERE BoxID < 16;")
        tempNameList = cursor.fetchall()
        nameList = []
        for name in tempNameList:
            nameList.append(name[0])
    
        # 'tagString' will be the output of this function
        tagString = ""
        # get the 'tagBytes' in the correct format for the following if statements
        tagBytes = str(tagBytes)
        tagBytes = tagBytes[2:-1]
    
        # Go through different sections of the 'tagBytes'
        # If a section == \x01\x01 then that Tag has been checked
        # and we can get that Tags name from the 'nameList' and add that name to 'tagString'
        if tagBytes[:8] == "\\x01\\x01":
            if 1 in boxIDList:
                tagString = tagString + nameList[0] +";"
        if tagBytes[8:16] == "\\x01\\x01":
            if 2 in boxIDList:
                tagString = tagString + nameList[1] +";"
        if tagBytes[16:24] == "\\x01\\x01":
            if 3 in boxIDList:
                tagString = tagString + nameList[2] +";"
        if tagBytes[24:32] == "\\x01\\x01":
            if 4 in boxIDList:
                tagString = tagString + nameList[3] +";"
        if tagBytes[32:40] == "\\x01\\x01":
            if 5 in boxIDList:
                tagString = tagString + nameList[4] +";"
        if tagBytes[40:48] == "\\x01\\x01":
            if 6 in boxIDList:
                tagString = tagString + nameList[5] +";"
        if tagBytes[48:56] == "\\x01\\x01":
            if 7 in boxIDList:
                tagString = tagString + nameList[6] +";"
        if tagBytes[56:64] == "\\x01\\x01":
            if 8 in boxIDList:
                tagString = tagString + nameList[7] +";"
        if tagBytes[64:72] == "\\x01\\x01":
            if 9 in boxIDList:
                tagString = tagString + nameList[8] +";"
        if tagBytes[72:80] == "\\x01\\x01":
            if 10 in boxIDList:
                tagString = tagString + nameList[9] +";"
        if tagBytes[80:88] == "\\x01\\x01":
            if 11 in boxIDList:
                tagString = tagString + nameList[10] +";"
        if tagBytes[88:96] == "\\x01\\x01":
            if 12 in boxIDList:
                tagString = tagString + nameList[11] +";"
        if tagBytes[96:104] == "\\x01\\x01":
            if 13 in boxIDList:
                tagString = tagString + nameList[12] +";"
        if tagBytes[104:112] == "\\x01\\x01":
            if 14 in boxIDList:
                tagString = tagString + nameList[13] +";"
        if tagBytes[112:120] == "\\x01\\x01":
            if 15 in boxIDList:
                tagString = tagString + nameList[14] +";"
        # Remove the last ;
        if len(tagString) > 0:
            tagString = tagString[:-1]
        
        return tagString
    
    # Get info about other errors
    except Exception as e:
        logger.error(e, exc_info=True)
        return -1


#####################################################################################
## Function: getColNameTupleListUpdatingInCD()
#####################################################################################
'''
This function gets a list of tuples.
Each tuple will contain a column DB name and a column Display name.
This function will get the correct columns for updating the CD Results file
'cdResultsFilePath' = the path to a CD results file
'cursor' = an SQLite cursor
'excelData' = a DataFrame containing Excel data
'excelFilePath' = the path to an Excel file
'oupdateColNameList' = a list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file)
'hideMessages' = Boolean value that controls whether or not the print statements get run. 
    Set to True to hide messages and return a list of outputs (default is False)
'''

# Gets the columns that are new, or editable
def getColNameTupleListUpdatingInCD(cdResultsFilePath, cursor, excelData, excelFilePath, updateColNameList, hideMessages):
    try:
        # Get the ID of compound table
        cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
        compoundTblID = cursor.fetchall()[0][0] 
        
        # This list is going to hold tuples, 
        # those tuples are going to hold 'colDBName' and 'colDisplayName'
        # This list is the output for this function
        colNameTupleList = []
       
        # If the user has chosen columns to update
        if updateColNameList is not None:
            # Loop through the columns that the user wants to update
            # 'updateColNameList' should be a list of the column display names
            for colDisplayName in updateColNameList:
                # If the column is in the Excel file
                if colDisplayName in excelData.columns:
                        
                    # If the column is in the CD results file
                    cursor.execute("SELECT COUNT(*) FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                    if cursor.fetchall()[0][0] != 0:
                    
                        # If the column is not 'Tags', we're making this check because 'Tags' is the byte type, 
                        # but we are still allowing the user to edit it
                        if colDisplayName != "Tags":
                            
                            # If the column is not stored in the CD results file as the bytes type
                            cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                            if cursor.fetchall()[0][0] != 6:
     
                                # If the column is editable
                                cursor.execute("SELECT Grid_AllowEdit FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                                if cursor.fetchall()[0][0] == 1:
                                        
                                    # Get the column DBName and Display name, put those names in a tuple, then add that tuple to the list of tuples
                                    cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                                    colNameTuple = cursor.fetchall()[0]
                                    colNameTupleList.append(colNameTuple)
                                        
                                # If the column is non-editable
                                else:
                                    print("WARNING: "+colDisplayName+" can't be updated because it is a non-editable column")
                            
                            # If the column is stored in the CD results file as the bytes type
                            else:
                                print("WARNING "+colDisplayName+" can't be updated because of the way this column is stored in the database")
                            
                        # If the column is the 'Tags' column
                        else:
                            # Get the column DBName and Display name, put those names in a tuple, then add that tuple to the list of tuples
                            cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                            colNameTuple = cursor.fetchall()[0]
                            colNameTupleList.append(colNameTuple)
                            
                    # If the column is not in the CD results file
                    else:
                        # convert the display name to a string that can be stored as an SQLite column name
                        colDBName = formatStringToSQLiteColumn(colDisplayName)
                                
                        # If the DB version of the column name is not already being used in the CD results file compound table
                        cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name=(?);", (colDBName, ))
                        if cursor.fetchall()[0][0] == 0:
                                    
                            colNameTuple = (colDBName, colDisplayName)
                            colNameTupleList.append(colNameTuple)
                                    
                        # If the DB version of the column name is already being used in the CD results file
                        else:
                            print("WARNING: can't add "+colDisplayName+" to "+cdResultsFilePath+" because the SQLite friendly version of the name ("+colDBName+") is already being used")
                       
                # If the column is not in the Excel file
                else:
                    print("WARNING: "+colDisplayName+" can't be found in "+excelFilePath)
        
        # If the user just wants to get all editable columns
        else:
            # Gets column DB names and display names from editable columns (Tags, Checked, Name, and columns added by user)
            # tempColNameTupleList will be a list of tuples
            cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Grid_AllowEdit = 1;", (compoundTblID, ))
            tempColNameTupleList = cursor.fetchall()
                    
            # Loop through each tuple in the temporary tuple list
            for colNameTuple in tempColNameTupleList:
                # If the column display name is an Excel file column name 
                if colNameTuple[1] in excelData.columns:
                    # Add tuple to the tuple list that will be used as output from this function
                    colNameTupleList.append(colNameTuple)
        
        return colNameTupleList
    
    # Get info about other errors
    except Exception as e:
        print("WARNING: Error")
        logger.error(e, exc_info=True)
        # Return an empty dataframe
        return None
        

#####################################################################################
## Function: getColNameTupleListUpdatingInExcel()
#####################################################################################
'''
This function gets a list of tuples.
Each tuple will contain a column DB name and a column Display name.
This function will get the correct columns for updating the Excel file
'cdResultsFilePath' = the path to a CD results file
'cursor' = an SQLite cursor
'excelData' = a DataFrame containing Excel data
'updateColNameList' = a list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file)
'hideMessages' = Boolean value that controls whether or not the print statements get run. 
    Set to True to hide messages and return a list of outputs (default is False)
'''

# Gets the columns that are new, or editable
def getColNameTupleListUpdatingInExcel(cdResultsFilePath, cursor, excelData, updateColNameList, hideMessages):
    try:            
        # Get the ID of compound table
        cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
        compoundTblID = cursor.fetchall()[0][0] 
        
        # This list is going to hold tuples, 
        # those tuples are going to hold colDBName and colDisplayName
        # This list is the output for this function
        colNameTupleList = []
        
        # If the user has chosen columns to update
        if updateColNameList is not None:
            # Loop through the columns that the user wants to update
            # updateColNameList should be a list of the column display names
            for colDisplayName in updateColNameList:
                    
                # If the column is in the CD results file
                cursor.execute("SELECT COUNT(*) FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                if cursor.fetchall()[0][0] != 0:
                    
                    # If the column is not 'Tags', we're making this check because 'Tags' is the byte type, 
                    # but we are still allowing the user to edit it
                    if colDisplayName != "Tags":
                            
                        # If the column is not stored in the CD results file as the bytes type
                        cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                        if cursor.fetchall()[0][0] != 6:
     
                            # Get the column DBName and Display name, put those names in a tuple, then add that tuple to the list of tuples
                            cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                            colNameTuple = cursor.fetchall()[0]
                            colNameTupleList.append(colNameTuple)
                            
                        # If the column is stored in the CD results file as the bytes type
                        else:
                            print("WARNING "+colDisplayName+" can't be updated because of the way this column is stored in the database")
                            
                    # If the column is the 'Tags' column
                    else:
                        # Get the column DBName and Display name, put those names in a tuple, then add that tuple to the list of tuples
                        cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                        colNameTuple = cursor.fetchall()[0]
                        colNameTupleList.append(colNameTuple)
                            
                # If the column is not in the CD results file
                else:
                    print("WARNING: "+colDisplayName+" can't be found in "+cdResultsFilePath)
 
        # If the user just wants to get all editable columns
        else:
            # Gets column DB names and display names from editable columns (Tags, Checked, Name, and columns added by user)
            # tempColNameTupleList will be a list of tuples
            cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Grid_AllowEdit = 1;", (compoundTblID, ))
            tempColNameTupleList = cursor.fetchall()
                    
            # Loop through each tuple in the temporary tuple list
            for colNameTuple in tempColNameTupleList:
                # If the column display name is an Excel file column name 
                if colNameTuple[1] in excelData.columns:
                    # Add tuple to the tuple list that will be used as output from this function
                    colNameTupleList.append(colNameTuple)
        
        return colNameTupleList
    
    # Get info about other errors
    except Exception as e:
        print("WARNING: Error")
        logger.error(e, exc_info=True)
        # Return an empty dataframe
        return None
        
        
#####################################################################################
## Function: fillNAValuesInDF()
#####################################################################################
'''
This function fills the NA values in a dataframe. 
If a column is not the Tags column, then empty rows are filled with default values
'excelData' = the DataFrame containing Excel data
'''

def fillNAValuesInDF(excelData):
    try:
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
    
    # Get info about other errors
    except Exception as e:
        logger.error(e, exc_info=True)


#####################################################################################
## Function: updateDataInCDResultsFile()
#####################################################################################
'''
    This function imports data from an Excel file to a Compound Discoverer (CD) results file.
    'cdResultsFilePath' = The path to a CD results file
    'excelFilePath' = The path to an Excel file
    'updateColNameList' = a list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
        all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file)
    'optionalTagList' = al list of binary column in the Excel file that the user wishes to use to set Tags in the CD results file (default in None),
        if this value is left as None, the Tag column can still be updated in the CD results file using the Tags string column in the Excel file
    'hideMessages' = Boolean value that controls whether or not the print statements get run. 
        Set to True to hide messages and return a list of outputs (default is False)
'''
    
def updateDataInCDResultsFile(cdResultsFilePath, excelFilePath, excelSheetName, updateColNameList = None, optionalTagList = None, hideMessages = False):
    cdResultsFileName = cdResultsFilePath.rpartition("/")[2]
    excelFileName = excelFilePath.rpartition("/")[2]
    
    if hideMessages == False: 
        print("\nImporting data from "+excelFileName+" into "+cdResultsFileName)
   
    try: 
        # If the results file can be found
        if os.path.exists(cdResultsFilePath):
        
            # Open connection to the Compound Discoverer File
            conn = sqlite3.connect(cdResultsFilePath)
            cursor = conn.cursor()
            
            excelData = pd.read_excel(excelFilePath, sheet_name = excelSheetName)
            excelData = fillNAValuesInDF(excelData)
            excelRowCount = len(excelData.index)
    
            # If the user wants to update the Tags in CD using boolean values from chosen Excel columns
            if optionalTagList is not None:
                # Change the names of the Tags in CD based on the Excel column the user has chosen with 'optionalTagList'
                # Also get the user selected tags that are boolean columns in the Excel file
                optionalTagList = changeTagNamesInCD(cursor, excelData, excelFilePath, optionalTagList, hideMessages)
    
            # Get list of tuples
            # Each tuple will contains the column DB name and display name
            colNameTupleList = getColNameTupleListUpdatingInCD(cdResultsFilePath, cursor, excelData, excelFilePath, updateColNameList, hideMessages)

            # If there was an error during the getColNameTupleList() function
            if colNameTupleList is None:
                raise Exception

            # Get the ID of compound table
            cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
            compoundTblID = cursor.fetchall()[0][0]
        
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
                                ((?), 'Cleaned', 4, 1, '3245F562-3044-4BC0-9091-3813CA7AE5BC', \
                                0, -1, '', 'Cleaned', 'Shows the rows that have been updated with CDExcelMessenger.py', \
                                '', 1, '',\
                                4, 0, -1, \
                                '', 0, 0);", (compoundTblID, ))
                # Set Cleaned to False for all rows
                cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET Cleaned = 'False';") 
            # If the 'Cleaned' column does exist
            else:
                # Check if 'ConsolidatedUnknownCompoundItems' has the 'OldCleaned' column
                cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='OldCleaned';")
                # If the 'OldCleaned' column doesn't exist
                if cursor.fetchall()[0][0] == 0:
                    # Rename 'Cleaned' to 'OldCleaned'
                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems RENAME COLUMN Cleaned to OldCleaned")
                    # Add Old Cleaned details to columns table
                    cursor.execute("INSERT INTO DataTypesColumns \
                                    (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType, \
                                    Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                                    Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                                    Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                                    Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                                    VALUES \
                                    ((?), 'OldCleaned', 4, 1, '3245F562-3044-4BC0-9091-3813CA7AE5BC', \
                                    0, -1, '', 'Old Cleaned', 'Shows the rows that have been updated with CDExcelMessenger.py', \
                                    '', 1, '',\
                                    4, 0, -1, \
                                    '', 0, 0);", (compoundTblID, ))
                
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

            # Create INDEX which might improve performance
            cursor.execute("CREATE INDEX IF NOT EXISTS MW_RT ON ConsolidatedUnknownCompoundItems (MolecularWeight, RetentionTime);")
            # Loop through each column tuple in the list of tuples, to update each column in the list
            for colNameTuple in colNameTupleList:

                colDBName = colNameTuple[0]
                colDisplayName = colNameTuple[1]

                # If the column doesn't exist in the CD results file, we need to add the column before updating it
                cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name=(?);", (colDBName, ))
                if cursor.fetchall()[0][0] == 0:
                
                    # Get the data type of the column
                    colDataType = excelData.dtypes[colDisplayName]
                           
                    # Add new column and set default values based on the data type of the columns
                    # Also set customDataType and valueType which gets used in the CD database
                    if colDataType == "int64":
                        # Add column to compound table
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" INTEGER;")
                        # Add default values to the new column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", (str(0), ))
                                
                        customDataType = "2"
                        valueType = "A170C73A-BD79-493B-B24A-B981BAF6DCC5"
                        #IF THIS DOESN'T WORK, TRY customDataType 1 and valueType B186F5FB-FD41-4087-8A54-CB52CCA0E3DF
                                
                    elif colDataType == "float64":
                        # Add column to compound table
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" REAL;")
                        # Add default values to the new column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", (colDBName, str(0.0), ))
                                
                        customDataType = "3"
                        valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                    elif colDataType == "bool":
                        # Add column to compound table
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" NUMERIC;")
                        # Add default values to the new column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", ("False", ))
                                
                        customDataType = "4"
                        valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                    else:
                        # Add column to compound table
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" TEXT;")
                        # Add default values to the new column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", ("", ))
                                
                        customDataType = "4"
                        valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                            
                    # Add column details to columns table
                    cursor.execute("INSERT INTO DataTypesColumns \
                                    (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType, \
                                    Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                                    Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                                    Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                                    Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                                    VALUES \
                                    ((?), (?), (?), 1, (?), \
                                    0, -1, '', (?), (?), \
                                    '', 1, '',\
                                    4, 0, -1, \
                                    '', 1, 0);", (compoundTblID, colDBName, customDataType, valueType, colDisplayName, colDisplayName+": This column has been added by CDExcelMessenger.py", ))
                    if hideMessages == False: 
                        print("Column: "+colDisplayName+" added to "+cdResultsFileName)
                    else:
                        outputList.append("Column: "+colDisplayName+" added to "+cdResultsFileName)

                # Loop through each row in the Excel file, to update the current column in the CD results file
                for row in range(excelRowCount):
                    # Molecular weight and retention time is needed to match rows between the excel file and CD results file
                    MW = excelData.at[row,"Calc. MW"]
                    RT = excelData.at[row,"RT [min]"]
                
                    # The Tags column needs to be handled differently
                    if colDBName == "Tags":
                        # If the user wants to update the Tags in CD using boolean values from chosen Excel columns
                        if optionalTagList is not None:
                        
                            # Create a tagString containing Tags that are checked for the current row
                            tagString = "" 
                            for optionalTag in optionalTagList:
                                boolVal = excelData.at[row,optionalTag]
                              
                                if boolVal == True:
                                    tagString = tagString + optionalTag + ";"
                                    
                            # Remove the last ;
                            if len(tagString) > 0:
                                tagString = tagString[:-1]
                            
                            # Convert the tag string to bytes that CD can read
                            value = getTagBytes(tagString, cdResultsFilePath, cursor)
                        
                        # If the user wants to update the Tags in CD with the Tags column from Excel 
                        else:
                            # Convert the tag string to bytes that CD can read
                            value = getTagBytes(excelData.at[row,colDisplayName], cdResultsFilePath, cursor)
                            if value == -1:
                                value = b"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"
      
                        # Update the current row and column in the CD results file, also set Cleaned to True
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?), Cleaned = 'True' WHERE ROUND(MolecularWeight, 5) = ROUND((?), 5) and ROUND(RetentionTime, 3) = ROUND((?), 3);", (value, str(MW), str(RT), ))     
                                    
                    # If the column isn't the Tags column
                    else:
                        # The Checked column also needs to be handled differently
                        if colDBName == "Checked": 
                            value = excelData.at[row,colDisplayName]
                            if bool(value):
                                value = 1
                            else:
                                value = 0
                        # Not the Tags or Checked column
                        else:
                            value = excelData.at[row,colDisplayName]
                            
                        # Update the current row and column in the CD results file, also set Cleaned to True
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?), Cleaned = 'True' WHERE ROUND(MolecularWeight, 5) = ROUND((?), 5) and ROUND(RetentionTime, 3) = ROUND((?), 3);", (str(value), str(MW), str(RT), ))     

                if hideMessages == False: 
                    print("Column: "+colDisplayName+" updated")
                else:
                    outputList.append("Column: "+colDisplayName+" updated")
            
            # Drop the INDEX that was created earlier    
            cursor.execute("DROP INDEX IF EXISTS MW_RT;") 
            conn.commit()
            if hideMessages == False: 
                print(cdResultsFileName+" updated")
        
            # Close the connection to the Compound Discoverer file
            cursor.close()
            conn.close()
        
            if hideMessages == True:
                return outputList      
    
       # If the results file can't be found
        else:
            print("WARNING: "+cdResultsFilePath+" can't be found")
        
    # Value Error can be caused by the Excel file not having the correct column names
    except ValueError:
        print("\nValue Error")
        print("Make sure "+excelFilePath+" has "+colName)
        
    # If the Excel file can't be found
    except FileNotFoundError:
        print("\nFileNotFoundError")
        print("Can't find "+excelFilePath)
    
    # If permission to the Excel file was denied
    except PermissionError:
        print("\nPermissionError")
        print("Couldn't gain permission to the Excel File")
        print("Make sure "+excelFilePath+" is not open in another program")
        
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
    
def updateDataInExcelFile(cdResultsFilePath, excelFilePath, excelSheetName, updateColNameList = None, optionalTagList = None, hideMessages = False):
    cdResultsFileName = cdResultsFilePath.rpartition("/")[2]
    excelFileName = excelFilePath.rpartition("/")[2]
    
    if hideMessages == False: 
        print("\nImporting data from "+cdResultsFileName+" into "+excelFileName)
    
    try: 
        # If the results file can be found
        if os.path.exists(cdResultsFilePath):
        
            # Open connection to the Compound Discoverer File
            conn = sqlite3.connect(cdResultsFilePath)
            cursor = conn.cursor()
        
            excelData = pd.read_excel(excelFilePath, sheet_name = excelSheetName)
            excelData = fillNAValuesInDF(excelData)
            excelRowCount = len(excelData.index)
            
            
            # Loop through each optional tag and append use a temp list
            # to get the tags that are boolean columns in the Excel file
            # and store those tags back into 'optionalTagList'
            tempOptionalTagList = []
            for optionalTag in optionalTagList:
                if optionalTag in excelData.columns:
                    if excelData.dtypes[optionalTag] == "bool":
                        tempOptionalTagList.append(optionalTag)
                    else:
                        print("WARNING: "+optionalTag+" isn't a boolean column")
                else:
                    print("WARNING: "+optionalTag+" can't be found in "+excelFilePath)
            if tempOptionalTagList != []:
                optionalTagList = tempOptionalTagList
            else:
                optionalTagList = None
            
        
            # Get list of tuples
            # Each tuple will contains the column DB name and display name
            colNameTupleList = getColNameTupleListUpdatingInExcel(cdResultsFilePath, cursor, excelData, updateColNameList, hideMessages)

            # If there was an error during the getColNameTupleList() function
            if colNameTupleList is None:
                raise Exception

            # Get the ID of compound table
            cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
            compoundTblID = cursor.fetchall()[0][0]

            # Loop through each column tuple in the list of tuples, to update each column in the list
            for colNameTuple in colNameTupleList:
                colDBName = colNameTuple[0]
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
                    cursor.execute("SELECT "+colDBName+" FROM ConsolidatedUnknownCompoundItems WHERE ROUND(MolecularWeight, 5) = ROUND((?), 5) and ROUND(RetentionTime, 3) = ROUND((?), 3);", (str(MW), str(RT), ))     
                    selectStatementResults = cursor.fetchall()
                    value = selectStatementResults[0][0]
                    
                    # Raise an exception if multiple rows in the CD results file matched with a row in the Excel file
                    if len(selectStatementResults) > 1:
                        if hideMessages == False: 
                            print("WARNING: multiple rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                            raise Exception
                            
                    # If exactly one row in the CD results file matched with a row in the Excel file
                    elif len(selectStatementResults) == 1:
                        # The Tags column needs to be handled differently
                        if colDBName == "Tags":
                            # Update the current row and column in the Excel data frame after converting the bytes value to a string
                            tagString = getTagString(value, cdResultsFilePath, cursor)
                            excelData.at[row,colDisplayName] = tagString
                            if optionalTagList is not None:
                            
                                # Create a list of tags, ';' is the delimiter
                                tagList = tagString.split(";")
                                
                                # Loop through the user chosen tags to update
                                for optionalTag in optionalTagList:
                                    # If the current tag is in 'tagList', 
                                    # that means the current tag for this row is checked in the CD results file
                                    if optionalTag in tagList:
                                        excelData.at[row,optionalTag]=True
                                    else:
                                        excelData.at[row,optionalTag]=False
                                     
                                
                        else:
                            # if the current column is a boolean column, but not Tags
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
                    
                    # No rows in the CD results file match with the current row in the Excel file
                    else:
                        if hideMessages == False: 
                            print("WARNING: no rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                        else:
                            outputList.append("WARNING: no rows in "+cdResultsFileName+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                # This is to make sure boolean values are set as bool in the Excel file
                if colIsBool:                 
                    excelData[colDisplayName] = excelData[colDisplayName].astype('bool')
            
            excelData.to_excel(excelFilePath, sheet_name="Peak", index=False)  
        
            if hideMessages == False: 
                print(excelFileName+" updated")
        

            # Close the connection to the Compound Discoverer file
            cursor.close()
            conn.close()
        
            if hideMessages == True:
                return outputList 
        
        # If the results file can't be found
        else:
            print("WARNING: "+cdResultsFilePath+" can't be found")

    # Value Error can be caused by the Excel file not having the correct column names
    except ValueError:
        print("\nValue Error")
        print("Make sure "+excelFilePath+" has "+colName)
        
    # If the Excel file can't be found
    except FileNotFoundError:
        print("\nFileNotFoundError")
        print("Can't find "+excelFilePath)
    
    # If permission to the Excel file was denied
    except PermissionError:
        print("\nPermissionError")
        print("Couldn't gain permission to the Excel File")
        print("Make sure "+excelFilePath+" is not open in another program")
    
    # Get info about errors
    except Exception as e:
        if hideMessages == False: 
            logger.error(e, exc_info=True)
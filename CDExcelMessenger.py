#####################################################################################
## Import Modules
#####################################################################################

import sqlite3
import os.path
import re
import pandas as pd
   

#####################################################################################
## Function: formatStringToSQLiteColumn()
#####################################################################################  
'''
This function converts a string into a string that can be used as an SQLite column name.
This is done by removing characters that aren't alphanumeric or underscores, then by
removing any numbers at the start of the column.

INPUT:
'string' = The string that needs to be converted to an SQLite friendly format.

OUTPUT:
'string' = The string after it has been converted to an SQLite friendly format.
''' 
    
def formatStringToSQLiteColumn(string):

    # Strip characters that aren't alphanumeric or underscores
    string = re.sub(r'\W+', '', string)
        
    # Strip numbers at the start of the column
    string = string.lstrip("0123456789") 
    
    return string
    

#####################################################################################
## Function: getValidTagNames()
#####################################################################################  
'''
This function gets two lists of valid tag names and makes sure there isn't more
than 15 tags. One list will contain the unique Tags only found in the Tags column 
of the Excel file (if the user included 'Tags' in the tagList),
the other list will contain the tags the user included in the tagList if 
those tags are valid.

INPUT:
'peakTable' = The peak table DataFrame.
'excelFilePath' = The path to the Excel file.
'tagList' = The list of names the user wishes to use as Tags, if they included "Tags" in this list,
    they want to include tags found in the 'Tags' column of the Excel file.

OUTPUT:
'tagList' = A list of valid tags, not including the tags that were only found in the 'Tags' 
    column of the Excel file.
'tagsInTagsCol' = A list of unique tags that were only found in the 'Tags' column of the Excel file.
'report' = A list of messages that can be printed to console if 'verbose' is True.
'''    
    
def getValidTagNames(peakTable, excelFilePath, tagList):
    try:
        report = []
        tempTagList = []
        tagsInTagsCol = []
    
        # If the user hasn't provided a tag list, or they have and "Tags" is in that list
        if tagList is None or (tagList is not None and "Tags" in tagList):
            tagSet = set()
            
            # If the Tags column in Excel contains strings
            if peakTable.dtypes["Tags"] == "object":

                for row in peakTable.index:
                    # If the current cell is not null
                    if pd.notnull(peakTable.at[row,"Tags"]):
                        # Get the string in the cell, split the string, remove whitespace from both ends of each tag, then add each tag to a set 
                        excelTag = peakTable.at[row, "Tags"]
                        excelTagSplit = excelTag.split(";")
                        
                        for tag in excelTagSplit:
                            tag = tag.strip()
                            tagSet.add(tag)
                            
            # If the Tags column in the Excel file does not contain strings
            else:
                report.append("WARNING: the 'Tags' column in "+excelFilePath+" doesn't contain strings")
        
            # tagsInTagsCol is going to contain the tags only found in the Excel Tags column
            tagsInTagsCol = list(tagSet)
        
        # If the user has provided a tag list
        if tagList is not None:
        
            # Get the binary columns in the Excel file
            binaryCols = peakTable.columns[peakTable.isin([0,1]).all()]
            
            for tag in tagList:
                if tag != "Tags":
                    # Keep track of whether or not the Tag is valid
                    validTag = False
                    
                    # If the tag is not a column in the Excel file,
                    if tag not in peakTable.columns:
                        validTag = True
                    
                    # If the tag is a column in the Excel file
                    else:
            
                        # If the column is not a boolean or integer column
                        if peakTable.dtypes[tag] != "bool" and peakTable.dtypes[tag] != "int64":
                            report.append("WARNING: the tag '"+tag+"' is a column in "+excelFilePath+", but the column is not boolean/binary")
                        
                        elif peakTable.dtypes[tag] == "int64":
                            # The column is an integer column, but not binary (1/0)
                            if tag not in binaryCols:
                                report.append("WARNING: the tag '"+tag+"' is a column in "+excelFilePath+", but the column is not boolean/binary")
                            
                            # The column is binary (1/0)
                            else:
                                validTag = True
                    
                        # Column is boolean
                        else:
                            validTag = True
                
                    if validTag:
                        if tag not in tempTagList:
                            tempTagList.append(tag)
                            # Remove from the list of tags that were found in the Excel Tags column
                            if tag in tagsInTagsCol:
                                tagsInTagsCol.remove(tag)
                                
                        # If the tag is already in the temporary tag list,
                        else:
                            report.append("WARNING: can't use the tag '"+tag+"' more than once")
                            
        tagList = tempTagList
        
        # Make sure there are 15 or less tags
        tagNum = len(tagList) + len(tagsInTagsCol)
        if tagNum > 15:
            raise ValueError("ValueError", "trying to add/update "+str(tagNum)+" Tags, the maximum number of Tags is 15")
        
        return tagList, tagsInTagsCol, report
            
    except Exception as e:
        raise e
    
   
#####################################################################################
## Function: changeCDTagsAndVisibility()
#####################################################################################
'''
This function changes the Tag names and visibility in the CD results file

INPUT:
'cursor' = An SQLite cursor.
'peakTable' = A DataFrame containing Excel data.
'cdResultsFilePath' = The path to a CD results file.
'tagList' = A list of strings the user wants to be tags.
'tagsInTagsCol' = A list of string of the unique tags only found in the Tags column of the Excel file.
'verbose' = Boolean value that controls the output to the console. 
    If False, hide output and return the output as a list.
'''

def changeCDTagsAndVisibility(cursor, peakTable, cdResultsFilePath, tagList, tagsInTagsCol, verbose):
    try:
        # Get IDs of all Tags in CD
        cursor.execute("SELECT BoxID FROM DataDistributionBoxExtendedData WHERE name = 'EntityItemTagVisibility';")
        tagIDList = []
        for tagID in cursor.fetchall():
            tagIDList.append(tagID[0])
            
        tempTagList = tagsInTagsCol + tagList
    
        # index will be used to access the IDs in the 'tagIDList'
        index = 0
        
        # Loop through each tag the user is adding/updating
        for tag in tempTagList:
            currID = tagIDList[index]
            
            # change Tag Name and Description in CD database
            cursor.execute("UPDATE DataDistributionBoxes SET Name = (?), Description = (?) WHERE BoxID = "+str(currID)+";", (tag, "Matching entry in: "+tag+".")) 
                    
            # Set the tags visibility to True
            cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'True' WHERE BoxID = "+str(currID)+";") 
                    
            index = index + 1
        
        # Loop through the CD tags that didn't have their names changed in CD
        # This is to make sure the only visible tags are the tags the user wants to use
        while index < (len(tagIDList)):
            currID = tagIDList[index]
            
            # Set the tags visibility to False
            cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'False' WHERE BoxID = "+str(currID)+";")
            
            index = index + 1
        
        if verbose:
            print("Tag names and visibility updated in "+cdResultsFilePath)
    
    # Operational Error 
    except sqlite3.OperationalError:
        raise sqlite3.OperationalError("SQLite3:OperationalError", "It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
    
    # Other errors
    except Exception as e:
        raise e
        
    
#####################################################################################
## Function: tagStringToBytes()
#####################################################################################
'''
This function receives a tags string and converts that 
string into bytes to be stored in the Compound Discoverer (CD) results file.

INPUT:
'tagString' = The string that needs to be converted to bytes.
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.

OUTPUT:
'tagbytes' = The tag bytes after being converted from a string.
'''

def tagStringToBytes(tagString, cdResultsFilePath, cursor):
    try:
        # If there are any Tags
        if not pd.isna(tagString) and tagString != "":
            
            # Create a list of tags, ';' is the delimiter
            tagStringSplit = tagString.split(";")
            
            # Loop through the list of the tags that were in the tag string to remove whitespace from the left and right of each tag
            for i in range(len(tagStringSplit)):
                tagStringSplit[i] = tagStringSplit[i].strip()
            
            # 'keys' will contain Tags that were in the tag string
            # 'keys' will be used in an SQLite statement to get the tag IDs
            keys = ""
            for tag in tagStringSplit:
                keys = keys + "'"+str(tag)+"', "
            
            # Remove last ', ' from the keys
            keys = keys[:-2]
            
            # Get the IDs of the Tags in 'keys' (the Tags that were in the tag string) 
            cursor.execute("SELECT BoxID FROM DataDistributionBoxes WHERE Name IN ("+keys+");")
            checkedTagIDList = []
            for ID in cursor.fetchall():
                checkedTagIDList.append(ID[0])
                
            # Get the IDs of all tags
            cursor.execute("SELECT BoxID FROM DataDistributionBoxExtendedData WHERE name = 'EntityItemTagVisibility';")
            tagIDList = []
            for ID in cursor.fetchall():
                tagIDList.append(ID[0])
                
            # 'tagBytes' will be the output of this function
            tagBytes = b""
            
            # Create the tag bytes in the correct format for the CD results file
            # if an ID is found in the checkedTagIDList, add \x01\x01 to indicate that the Tag with that ID is checked
            # \x00\x00 is used to idicate that the Tag with that ID is not checked
            for ID in tagIDList:
                
                # If the current ID is not found in the ID list
                if ID not in checkedTagIDList:
                    tagBytes += b"\x00\x00"
                
                # If the current ID is found in the ID list
                else:
                    tagBytes += b"\x01\x01"
            
            return tagBytes
        
        # There is no tags
        else:
            return None
    
    # Operational Error 
    except sqlite3.OperationalError:
        raise sqlite3.OperationalError("SQLite3:OperationalError", "It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
    
    # Other errors
    except Exception as e:
        raise e
        
        
#####################################################################################
## Function: tagBytesToString()
#####################################################################################
'''
This function receives tag bytes and converts those 
bytes into a string to be stored in the Excel file if the Tag is set as visible. 

INPUT:
'tagBytes' = The bytes that needs to be converted to a string.
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.

OUTPUT:
'tagString' = The tag string after being converted from bytes.
'''

def tagBytesToString(tagBytes, cdResultsFilePath, cursor):
    try:    
        # Get a list of IDs for the Tags that are visible in CD
        cursor.execute("SELECT BoxID FROM DataDistributionBoxExtendedData WHERE name = 'EntityItemTagVisibility' AND ValueString = 'True';")
        visibleIDList = []
        for ID in cursor.fetchall():
            visibleIDList.append(ID[0])
        
        # Get the IDs of all tags
        cursor.execute("SELECT BoxID FROM DataDistributionBoxExtendedData WHERE name = 'EntityItemTagVisibility';")
        tagIDList = []
        for ID in cursor.fetchall():
            tagIDList.append(ID[0])        
            
        # 'keys' will contain IDs of all the tags
        # 'keys' will be used in an SQLite statement to get the tag Names
        keys = ""
        for ID in tagIDList:
            keys = keys + "'"+str(ID)+"', "
            
        # Remove last ', ' from the keys
        keys = keys[:-2]
            
        # Get the Names of the Tags in 'keys' (all the Tags) 
        cursor.execute("SELECT Name FROM DataDistributionBoxes WHERE BoxID IN ("+keys+");")
        tagNameList = []
        for tag in cursor.fetchall():
            tagNameList.append(tag[0])    
            
        # 'tagString' will be the output of this function
        tagString = ""
        
        # get the 'tagBytes' in the correct format for the following code
        tagBytes = str(tagBytes)
        tagBytes = tagBytes[2:-1]
    
        # Go through different sections of the 'tagBytes'
        # If a section == \x01\x01 then that Tag has been checked
        # and we can get the name for that tag from the 'tagNameList' and add that name to 'tagString'
        start = 0
        end = 8
        for i in range(len(tagIDList)):
        
            # If the current tag is checked in CD
            if tagBytes[start : end] == "\\x01\\x01":
            
                # If the current tag is visible in CD
                if tagIDList[i] in visibleIDList:
                    tagString = tagString + tagNameList[i] +";"
            
            start = start + 8
            end = end + 8
                
        # Remove the last ';'
        if len(tagString) > 0:
            tagString = tagString[:-1]
        
        return tagString
    
    # Operational Error 
    except sqlite3.OperationalError:
        raise sqlite3.OperationalError("SQLite3:OperationalError", "It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
    
    # Other errors
    except Exception as e:
        raise e


#####################################################################################
## Function: getColNamesForUpdatingCD()
#####################################################################################
'''
This function gets a list of tuples.
Each tuple will contain a column DB name and a column Display name.
This function will get the correct columns for updating the Compound Discoverer (CD) Results file.

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.
'peakTable' = A DataFrame containing Excel data.
'excelFilePath' = The path to an Excel file.
'excelColList' = A list of columns in the Excel file that the user wishes to update. If this value is None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has already added to the CD results file).
    
OUTPUT:
'colNameTupleList' = A list of tuples. Each tuple contains a column's DB name and display name.
'report' = A list of messages that can be printed to console if 'verbose' is True.
'''

# Gets the columns that are new, or editable
def getColNamesForUpdatingCD(cdResultsFilePath, cursor, peakTable, excelFilePath, excelColList):
    try:
        report = []
    
        # Get the ID of the compound table in CD
        cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
        compoundTblID = cursor.fetchall()[0][0] 
        
        # Get the custom data types and their IDs from CD, then store those values in a dictionary
        # The dictionary keys will be the data types and the dictionary values will be the IDs
        cursor.execute("SELECT Value, Name FROM CustomDataTypes;")
        cdDataTypeDict = {}
        for dataType in cursor.fetchall():
            cdDataTypeDict[dataType[1]] = dataType[0]
        
        # This list is going to hold tuples, 
        # those tuples are going to hold 'colDBName' and 'colDisplayName'
        # This list is the output for this function
        colNameTupleList = []
        
        # This is going to hold a list of column DB names,
        # The purpose of this list is to make sure we don't have multiple 
        # columns with the same DB name
        colDBNameList = []
       
        # If the user has chosen columns in the Excel file to use
        if excelColList is not None:
        
            # Loop through the columns that the user wants to update
            # 'excelColList' should be a list of the column display names
            for colDisplayName in excelColList:
                
                # If the column is in the Excel file
                if colDisplayName in peakTable.columns:
                        
                    # If the column is in the CD results file
                    cursor.execute("SELECT COUNT(*) FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                    if cursor.fetchall()[0][0] != 0:
                            
                        # If the column is not stored in the CD results file as bytes
                        cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                        if cursor.fetchall()[0][0] != cdDataTypeDict["Binary"]:
     
                            # If the column is editable
                            cursor.execute("SELECT Grid_AllowEdit FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                            if cursor.fetchall()[0][0] == 1:
                                
                                # Get the column DBName and Display name, put those names in a tuple
                                cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                                colNameTuple = cursor.fetchall()[0]
                                
                                # If the column is not the originalName column
                                if colNameTuple[0] != "originalName":
                                
                                    # If the column is not the MSI, Tags, or Notes column, 
                                    if colNameTuple[0] != "MSI" and colNameTuple[0] != "Tags" and colNameTuple[0] != "Notes":
                                        
                                        # Make sure we don't add multple columns with the same DB name
                                        if colNameTuple[0] not in colDBNameList:
                                            colNameTupleList.append(colNameTuple)
                                            colDBNameList.append(colNameTuple[0])
                                        else:
                                            report.append("WARNING: trying to add multiple columns with the SQLite name "+colNameTuple[0])
                                    
                                # If the column is the originalName column    
                                else:
                                    report.append("WARNING: "+colDisplayName+" can't be updated ")
                                    
                            # If the column is non-editable
                            else:
                                report.append("WARNING: "+colDisplayName+" can't be updated because it is a non-editable column")
                            
                        # If the column is stored in the CD results file as the bytes type
                        else:
                            report.append("WARNING "+colDisplayName+" can't be updated because of the way this column is stored in the database")
                            
                    # If the column is not in the CD results file
                    else:
                    
                        # convert the display name to a string that can be stored as an SQLite column name
                        colDBName = formatStringToSQLiteColumn(colDisplayName)
                                
                        # If the DB version of the column name is not already being used in the CD results file compound table
                        cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name=(?);", (colDBName, ))
                        if cursor.fetchall()[0][0] == 0:
                            
                            # Add DBName and Display name in a tuple
                            colNameTuple = (colDBName, colDisplayName)
                            
                            # If the column is not the originalName column
                            if colNameTuple[0] != "originalName":
                                
                                # If the column is not the MSI, Tags, or Notes column
                                if colNameTuple[0] != "MSI" and colNameTuple[0] != "Tags" and colNameTuple[0] != "Notes":
                                
                                    # Make sure we don't add multple columns with the same DB name
                                    if colNameTuple[0] not in colDBNameList:
                                        colNameTupleList.append(colNameTuple)
                                        colDBNameList.append(colNameTuple[0])
                                    else:
                                        report.append("WARNING: trying to add multiple columns with the SQLite name "+colNameTuple[0])
                            
                            # If the column is the originalName column    
                            else:
                                report.append("WARNING: "+colDisplayName+" can't be updated ")
                                    
                        # If the DB version of the column name is already being used in the CD results file
                        else:
                            report.append("WARNING: can't add "+colDisplayName+" to "+cdResultsFilePath+" because the SQLite friendly version of the name ("+colDBName+") is already being used")
                       
                # If the column is not in the Excel file
                else:
                    report.append("WARNING: "+colDisplayName+" can't be found in "+excelFilePath)
        
        # If the user just wants to get all editable columns
        else:
        
            # Gets column DB names and display names from editable columns (Tags, Checked, Name, and columns added by user)
            # tempColNameTupleList will be a list of tuples
            cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Grid_AllowEdit = 1;", (compoundTblID, ))
            tempColNameTupleList = cursor.fetchall()
                    
            # Loop through each tuple in the temporary tuple list
            for colNameTuple in tempColNameTupleList:
            
                # If the column display name is an Excel file column name 
                if colNameTuple[1] in peakTable.columns:
                    
                    # Make sure we don't get the originalName, MSI, Tags, or Notes columns here
                    if colNameTuple[0] != "originalName" and colNameTuple[0] != "MSI" and colNameTuple[0] != "Tags" and colNameTuple[0] != "Notes":
                
                        # Add tuple to the tuple list that will be used as output from this function
                        colNameTupleList.append(colNameTuple)
        
        # Add the Notes column with the correct display name
        colNameTupleList.append(("Notes", "Notes"))

        return colNameTupleList, report
    
    # Operational Error 
    except sqlite3.OperationalError:
        raise sqlite3.OperationalError("SQLite3:OperationalError", "It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
    
    # Other errors
    except Exception as e:
        raise e
        

#####################################################################################
## Function: getColNamesForUpdatingExcel()
#####################################################################################
'''
This function gets a list of tuples.
Each tuple will contain a column DB name and a column Display name.
This function will get the correct columns for updating the Excel file.

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.
'peakTable' = A DataFrame containing Excel data.
'excelColList' = A list of columns in the Excel file that the user wishes to update. If this value is None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file).
    
OUTPUT:
'colNameTupleList' = A list of tuples. Each tuple contains a column's DB name and display name.
'report' = A list of messages that can be printed to console if 'verbose' is True.
'''

# Gets the columns that are new, or editable
def getColNamesForUpdatingExcel(cdResultsFilePath, cursor, peakTable, excelColList):
    try: 
        report = []
    
        # Get the ID of compound table
        cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
        compoundTblID = cursor.fetchall()[0][0] 
        
        # Get the custom data types and their IDs from CD, then store those values in a dictionary
        # The dictionary keys will be the data types and the dictionary values will be the IDs
        cursor.execute("SELECT Value, Name FROM CustomDataTypes;")
        cdDataTypeDict = {}
        for dataType in cursor.fetchall():
            cdDataTypeDict[dataType[1]] = dataType[0]
        
        # This list is going to hold tuples, 
        # those tuples are going to hold colDBName and colDisplayName
        # This list is the output for this function
        colNameTupleList = []
        
        # If the user has chosen columns to update
        if excelColList is not None:
        
            # Loop through the columns that the user wants to update
            # excelColList should be a list of the column display names
            for colDisplayName in excelColList:
                    
                # If the column is in the CD results file
                cursor.execute("SELECT COUNT(*) FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                if cursor.fetchall()[0][0] != 0:
                        
                    # If the column is not stored in the CD results file as the bytes type, or is the Tags column
                    cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                    if cursor.fetchall()[0][0] != cdDataTypeDict["Binary"] or colDisplayName == "Tags":
     
                        # Get the column DBName and Display name, put those names in a tuple
                        cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                        colNameTuple = cursor.fetchall()[0]
                        
                        # Make sure MSI doesn't get added to the list
                        if colNameTuple[0] != "MSI":
                            colNameTupleList.append(colNameTuple)
                            
                    # If the column is stored in the CD results file as the bytes type
                    else:
                        report.append("WARNING "+colDisplayName+" can't be updated because of the way this column is stored in the database")
                 
                # If the column is not in the CD results file
                else:
                    report.append("WARNING: "+colDisplayName+" can't be found in "+cdResultsFilePath)
    
        # If the user just wants to get all editable columns
        else:
        
            # Gets column DB names and display names from editable columns (Tags, Checked, Name, and columns added by user)
            # tempColNameTupleList will be a list of tuples
            cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Grid_AllowEdit = 1;", (compoundTblID, ))
            tempColNameTupleList = cursor.fetchall()
                    
            # Loop through each tuple in the temporary tuple list
            for colNameTuple in tempColNameTupleList:

                # If the column display name is an Excel file column name 
                if colNameTuple[1] in peakTable.columns:

                    # Make sure MSI doesn't get added to the list
                    if colNameTuple[0] != "MSI":

                        # Add tuple to the tuple list that will be used as output from this function
                        colNameTupleList.append(colNameTuple)
        
        return colNameTupleList, report
    
    # Operational Error 
    except sqlite3.OperationalError:
        raise sqlite3.OperationalError("SQLite3:OperationalError", "It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
    
    # Other errors
    except Exception as e:
        raise e
        
        
#####################################################################################
## Function: fillNAValuesInDF()
#####################################################################################
'''
This function fills the NA values in a dataframe. 
If a column is not the Tags column, then empty rows are filled with default values.

INPUT:
'peakTable' = The DataFrame containing Excel data.

OUTPUT:
'peakTable' = The DataFrame containing Excel data after NA values have been filled.
'''

def fillNAValuesInDF(peakTable):
    for column in peakTable: 
        if column != "Tags":
            # If all values in the column are null
            if peakTable[column].isnull().all():
                peakTable[column] = peakTable[column].fillna("")
            else:
                # If the column has boolean values
                if peakTable.dtypes[column] == "bool":
                    peakTable[column] = peakTable[column].fillna(False)
                # If the column has int values
                elif peakTable.dtypes[column] == "int64":
                    peakTable[column] = peakTable[column].fillna(0)
                # If the column has float values
                elif peakTable.dtypes[column] == "float64":
                    peakTable[column] = peakTable[column].fillna(0.0)
                # If the column is not the boolean, integer, or float type
                else:
                    peakTable[column] = peakTable[column].fillna("")
   
    return peakTable


#####################################################################################
## Function: createCompoundIDColumns()
#####################################################################################
'''
This function adds the Compound Discoverer (CD) results file compound IDs to the Excel file.
Rows in the CD database and the Excel file are matched using the RetentionTime and MolecularWeight.
This is slow, but we only need to do this once. Matching rows will be faster once we have gotten the IDs.

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'conn' = The SQLite connection.
'cursor' = An SQLite cursor.
'peakRowCount' = The number of rows in the Excel Peak sheet.
'excelFilePath' = The path to an Excel file.
'peakTable' = The dataframe containing Excel data.
'peakSheetName' = The name of the Excel sheet containing the peak data.

OUTPUT:
'peakTable' = The dataframe containing the Excel data, now with IDs.
'report' = A list of messages that can be printed to console if 'verbose' is True.
'''

def createCompoundIDColumns(cdResultsFilePath, conn, cursor, peakRowCount, excelFilePath, peakTable, peakSheetName):
    try:
        report = []
    
        # list to hold the IDs that will get added to the Excel file
        IDList = []
            
        # If the compoundID column doesn't exist in CD, create it
        cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='compoundID';")
        if cursor.fetchall()[0][0] == 0:
            cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN compoundID")
            
            # Get the ID of the compound table in the CD database
            cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
            compoundTblID = cursor.fetchall()[0][0]
            
            # Get the custom data types and their IDs from CD, then store those values in a dictionary
            # The dictionary keys will be the data types and the dictionary values will be the IDs
            cursor.execute("SELECT Value, Name FROM CustomDataTypes;")
            cdDataTypeDict = {}
            for dataType in cursor.fetchall():
                cdDataTypeDict[dataType[1]] = dataType[0]
                
            # Add compoundID details to columns table
            cursor.execute("INSERT INTO DataTypesColumns \
                            (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType,\
                            Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                            Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                            Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                            Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                            VALUES \
                            ((?), 'compoundID', (?), 1, '3245F562-3044-4BC0-9091-3813CA7AE5BC', \
                            0, -1, '', 'compoundID', 'The database unique IDs. Matches with the Excel file.', \
                            '', 1, '',\
                            4, 0, -1, \
                            '', 0, 0);", (compoundTblID, str(cdDataTypeDict["String"]), ))
                                
            # Set Cleaned to False for all rows
            cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET compoundID = NULL;") 
            
        # Create a MolecularWeight & RetentionTime INDEX which improves performance with the upcoming SELECT statement
        cursor.execute("CREATE INDEX IF NOT EXISTS MW_RT ON ConsolidatedUnknownCompoundItems (MolecularWeight, RetentionTime);")
    
        # The MW and RT columns likely have one of these names
        if "Calc. MW" in peakTable.columns:
            mwName = "Calc. MW"
        elif "MW" in peakTable.columns:
            mwName = "MW"
        elif "MolecularWeight" in peakTable.columns:
            mwName = "MolecularWeight"
        else:
            raise ValueError("ValueError", "Can't find 'MW' column in the Excel file")
        if "RT [min]" in peakTable.columns:
            rtName = "RT [min]"
        elif "RT" in peakTable.columns:
            rtName = "RT"
        elif "RetentionTime" in peakTable.columns:
            rtName = "RetentionTime"
        else:
            raise ValueError("ValueError", "Can't find 'RT' column in the Excel file")
        
        # Make sure the MW and RT columns are float columns
        if peakTable.dtypes[mwName] != "float64":
            raise TypeError("TypeError", mwName+" is not a float column in the Excel file")
        
        if peakTable.dtypes[rtName] != "float64":
            raise TypeError("TypeError", rtName+" is not a float column in the Excel file")
        
        # Loop through each row in the excel file
        for row in range(peakRowCount):
            if pd.notnull(peakTable.at[row,mwName]) and pd.notnull(peakTable.at[row,rtName]):
                # Get molecular weight and retention time from the Excel data        
                # Molecular weight and retention time is needed to match rows between the excel file and CD results file
                MW = peakTable.at[row,mwName]
                RT = peakTable.at[row,rtName]
                        
                # Get the ID of the current row and column from the CD results file
                # Round MolecularWeight to 5 decimal places and RetentionTime to 3 decimal places because those values are stored in that format in the Excel file
                cursor.execute("SELECT ID FROM ConsolidatedUnknownCompoundItems WHERE ROUND(MolecularWeight, 5) = ROUND((?), 5) and ROUND(RetentionTime, 3) = ROUND((?), 3);", (str(MW), str(RT), ))     
                selectStatementResults = cursor.fetchall()
                        
                # If exactly one row in the CD results file matched with a row in the Excel file
                if len(selectStatementResults) == 1:
                    ID = selectStatementResults[0][0]
                    # Add current ID to the ID list which will be added to the Excel file
                    IDList.append(ID)
                
                    # Add the ID to the compoundID column in the CD results file
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET compoundID = (?) WHERE ID = (?);", (str(ID), str(ID), )) 
            
                # If multiple rows in the CD results file matched with a row in the Excel file
                elif len(selectStatementResults) > 1:
                    report.append("WARNING: multiple rows in "+cdResultsFilePath+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))

                # If no rows in the CD results file matched with a row in the Excel file
                else:
                    report.append("WARNING: no rows in "+cdResultsFilePath+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
            
            # One of the MW or RT cells was empty
            else:
                raise ValueError("ValueError", "At least one of the cells in the "+mwName+" or "+rtName+" columns are empty")
        
        # Drop the INDEX that was created earlier because it is not needed anymore   
        cursor.execute("DROP INDEX IF EXISTS MW_RT;")
                
        try:        
            # Update the peak dataframe with the ID list, then save the dataframe to the Excel file 
            IDData = pd.DataFrame (IDList, columns = ["compoundID"])
            peakTable = pd.concat([IDData, peakTable], axis=1)
            
            with pd.ExcelWriter(
                excelFilePath,
                mode="a",
                engine="openpyxl",
                if_sheet_exists="replace",
            ) as writer:
                peakTable.to_excel(writer, sheet_name=peakSheetName, index=False)  
        
        # If the Excel file doesn't have the correct sheet
        except ValueError:
            raise ValueError("ValueError", "Can't find "+peakSheetName+" in "+excelFilePath)
        
        # If the Excel file can't be found
        except FileNotFoundError:
            raise FileNotFoundError("FileNotFoundError", "Can't find "+excelFilePath)
    
        # If permission to the Excel file was denied
        except PermissionError:
            raise PermissionError("PermissionError", "Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")
                
        report.append("Column: compoundID added to "+excelFilePath)
        report.append("Column: compoundID added to "+cdResultsFilePath)
    
        conn.commit()
        return peakTable, report    
    
    # Operational Error 
    except sqlite3.OperationalError:
        raise sqlite3.OperationalError("SQLite3:OperationalError", "It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
    
    # Other errors
    except Exception as e:
        raise e
        
        
#####################################################################################
## Function: getMSIValue()
#####################################################################################
'''
This function calculates the MSI of a peak based on the Tags that have been checked.

INPUT: 
'tagList' = A list of tags that have been checked for the peak

OUTPUT:
'msi' = The MSI value of the peak that is based the tag list.
'''

def getMSIValue(tagList):
    msi = 0

    if "ms2Hit" in tagList:
        msi = msi + 1
    if "goodPeakShape" in tagList:
        msi = msi + 1
    if "goodRT" in tagList:
        msi = msi + 1
    if "queryMS" in tagList:
        msi = msi + 1
    if "msError" in tagList:
        msi = msi + 1
    
    return msi
     
       
        
#####################################################################################
## Function: validateUpdateCDInput()
#####################################################################################
'''
This function makes sure the user input for updating the CD results file is all valid

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'excelFilePath' = The path to an Excel file.
'peakSheetName' = The name of the Excel sheet containing the peak data.
'excelColList' = a list of columns in the Excel file that the user wishes to update.
'tagList' = The list of Tags that the user wishes to use.
'verbose' = Boolean value that controls the output to the console. 
'''

def validateUpdateCDInput(cdResultsFilePath, excelFilePath, peakSheetName, excelColList, tagList, verbose):
    # Validate 'cdResultsFilePath'
    if type(cdResultsFilePath) != str:
        raise TypeError("TypeError", "Make sure 'cdResultsFilePath' is a string value")
     
    # Validate 'excelFilePath'
    if type(excelFilePath) != str:
        raise TypeError("TypeError", "Make sure 'excelFilePath' is a string value")

    # Validate 'peakSheetName'
    if type(peakSheetName) != str:
        raise TypeError("TypeError", "Make sure 'peakSheetName' is a string value")
    
    # Validate 'excelColList'
    if excelColList is not None:
        if type(excelColList)!= list:
            raise TypeError("TypeError", "Make sure 'excelColList' is a list")
        
        for i in excelColList:
            if type(i) != str:
                raise TypeError("TypeError", "Make sure all values of 'excelColList' are string values")
    
    # Validate 'tagList'
    if tagList is not None:
        if type(tagList)!= list:
            raise TypeError("TypeError", "Make sure 'tagList' is a list")
        
        for i in tagList:
            if type(i) != str:
                raise TypeError("TypeError", "Make sure all values of 'tagList' are string values")
    
    # Validate 'verbose'
    if type(verbose) != bool:
        raise TypeError("TypeError", "Make sure 'verbose' is a boolean value")
    

#####################################################################################
## Function: updateCDResultsFile()
#####################################################################################
'''
This function imports data from an Excel file to a Compound Discoverer (CD) results file.

INPUT:
'cdResultsFilePath' = The path to a CD results file
'excelFilePath' = The path to an Excel file
'peakSheetName' = The name of the Excel sheet containing the peak data
'excelColList' = a list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file)
'tagList' = The list of Tags that the user wishes to use.
'verbose' = Boolean value that controls the output to the console. 
    If False, hide outputs and return the outputs as a list.
    
OUTPUT:    
'report' = A list of messages that can be printed to console if 'verbose' is True.
'''
    
def updateCDResultsFile(cdResultsFilePath, excelFilePath, peakSheetName, excelColList = None, tagList = None, verbose = True):
    # Set the sqlite connection and cursor values to None so that they can be closed during exceptions
    # only if the exception happened after the connections have been set
    conn = None
    cursor = None
    try:
        # Basic validation of user input
        validateUpdateCDInput(cdResultsFilePath, excelFilePath, peakSheetName, excelColList, tagList, verbose)
    
        report = []
        if verbose: 
            print("\nImporting data from "+excelFilePath+" into "+cdResultsFilePath)
    
        # If the results file can't be found
        if os.path.exists(cdResultsFilePath) == False:
            raise FileNotFoundError("FileNotFoundError", cdResultsFilePath+" can't be found")
          
        # Open connection to the Compound Discoverer File
        conn = sqlite3.connect(cdResultsFilePath)
        cursor = conn.cursor()
        
        # Get the custom data types and their IDs from CD, then store those values in a dictionary
        # The dictionary keys will be the data types and the dictionary values will be the IDs
        cursor.execute("SELECT Value, Name FROM CustomDataTypes;")
        cdDataTypeDict = {}
        for dataType in cursor.fetchall():
            cdDataTypeDict[dataType[1]] = dataType[0]
        
        try:
            # Get Excel data in a dataframe, fill NA values in that dataframe, and get the number of rows in that dataframe
            peakTable = pd.read_excel(excelFilePath, sheet_name = peakSheetName)
            peakTable = fillNAValuesInDF(peakTable)
            peakRowCount = len(peakTable.index)
            
        # If the Excel file doesn't have the correct sheet
        except ValueError:
            raise ValueError("ValueError", "Can't find "+peakSheetName+" in "+excelFilePath)
        
        # If the Excel file can't be found
        except FileNotFoundError:
            raise FileNotFoundError("FileNotFoundError", "Can't find "+excelFilePath)
    
        # If permission to the Excel file was denied
        except PermissionError:
            raise PermissionError("PermissionError", "Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")  

        # Get list of tuples
        # Each tuple will contains the column DB name and display name
        colNameTupleList, newReport = getColNamesForUpdatingCD(cdResultsFilePath, cursor, peakTable, excelFilePath, excelColList)
        for i in newReport:
            report.append(i)
        
        tagsInTagsCol = []
        # If the user has not chosen columns to update, or they have provided the tagList
        if excelColList is None or tagList is not None:
            # Get the valid Tags
            # 'tagsInTagsCol' is going to be a list of unique tags that are only found in the Tags column in the Excel file
            # 'tagList' is going to be a list of valid Tags that the user chose
            tagList, tagsInTagsCol, newReport = getValidTagNames(peakTable, excelFilePath, tagList)
            for i in newReport:
                report.append(i)
        
            # Change the names of the Tags in CD and set the visibility
            changeCDTagsAndVisibility(cursor, peakTable, cdResultsFilePath, tagList, tagsInTagsCol, verbose)
            
            # Add Tags tuple to list of tuples so the Tags column can be updated
            colNameTupleList.append(("Tags", "Tags"))
                
        # Get the ID of the compound table in the CD database
        cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
        compoundTblID = cursor.fetchall()[0][0]
        
        # If the 'Cleaned' column doesn't exist, create it
        cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='Cleaned';")
        if cursor.fetchall()[0][0] == 0:
            cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN Cleaned TEXT")
                
            # Add Cleaned details to columns table
            cursor.execute("INSERT INTO DataTypesColumns \
                            (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType,\
                            Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                            Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                            Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                            Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                            VALUES \
                            ((?), 'Cleaned', (?), 1, '3245F562-3044-4BC0-9091-3813CA7AE5BC', \
                            0, -1, '', 'Cleaned', 'Shows the rows that have been updated with CDExcelMessenger.py', \
                            '', 1, '',\
                            4, 0, -1, \
                            '', 0, 0);", (compoundTblID, str(cdDataTypeDict["String"]), ))
            
            report.append("Column: Cleaned added to "+cdResultsFilePath)
            
        # Set Cleaned to False for all rows
        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET Cleaned = 'False';") 
        
        # If the 'originalName' column doesn't exist, create it
        cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='originalName';")
        if cursor.fetchall()[0][0] == 0:
            cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN originalName TEXT")
                
            # Add originalName details to columns table
            cursor.execute("INSERT INTO DataTypesColumns \
                            (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType,\
                            Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                            Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                            Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                            Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                            VALUES \
                            ((?), 'originalName', (?), 1, '3245F562-3044-4BC0-9091-3813CA7AE5BC', \
                            0, -1, '', 'originalName', 'The original names in Compound Discoverer', \
                            '', 1, '',\
                            4, 0, -1, \
                            '', 0, 0);", (compoundTblID, str(cdDataTypeDict["String"]), ))
            
            report.append("Column: originalName added to "+cdResultsFilePath)
            
            # Set originalName values
            cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET originalName = Name;") 
        
        # If the 'MSI' column doesn't exist, create it
        cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='MSI';")
        if cursor.fetchall()[0][0] == 0:
            cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN MSI TEXT")
                
            # Add Cleaned details to columns table
            cursor.execute("INSERT INTO DataTypesColumns \
                            (DataTypeID, DBColumnName, CustomDataType, Nullable, ValueType,\
                            Creator, Finalizer, Property_Guid, Property_DisplayName, Property_Description, \
                            Property_FormatString, Property_SortDirection, Property_SemanticDescription, \
                            Grid_DataVisibility, Grid_VisiblePosition, Grid_ColumnWidth, \
                            Grid_GridCellControlGuid, Grid_AllowEdit, Grid_Background) \
                            VALUES \
                            ((?), 'MSI', (?), 1, '3245F562-3044-4BC0-9091-3813CA7AE5BC', \
                            0, -1, '', 'MSI', 'The MSI value calculated by CDExcelMessenger.py based on the Tags that were checked.', \
                            '', 1, '',\
                            4, 0, -1, \
                            '', 1, 0);", (compoundTblID, str(cdDataTypeDict["String"]), ))
            
            report.append("Column: MSI added to "+cdResultsFilePath)
            
        # Set MSI to NULL for all rows
        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET MSI = NULL;") 
        
        # If the Excel data doesn't contain the CD database IDs, add them
        if "compoundID" not in peakTable.columns:
            peakTable, newReport = createCompoundIDColumns(cdResultsFilePath, conn, cursor, peakRowCount, excelFilePath, peakTable, peakSheetName) 
            for i in newReport:
                report.append(i)

        # Loop through each column tuple in the list of tuples, to update each column in the list
        for colNameTuple in colNameTupleList:

            # Get the column DB name and display name
            colDBName = colNameTuple[0]
            colDisplayName = colNameTuple[1]

            # If the display name isn't a column in the Excel data, set default type
            if colDisplayName not in peakTable.columns:
                colDataType = "object"
                    
            # If the display name is a column in the Excel data
            else:
                # Get the data type of the Excel column
                colDataType = peakTable.dtypes[colDisplayName]

            # If the column doesn't exist in the CD results file, we need to add the column before updating it
            cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name=(?);", (colDBName, ))
            if cursor.fetchall()[0][0] == 0:
                           
                # Add new column to CD results file and set default values based on the data type of the columns
                # Also set 'customDataType' and 'valueType' which gets used in the CD database
                if colDataType == "int64":
                    # Add column to compound table
                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" INTEGER;")
                    # Add default values to the new column
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", (str(0), ))
                                
                    customDataType = str(cdDataTypeDict["Int64"])
                    valueType = "A170C73A-BD79-493B-B24A-B981BAF6DCC5"
                                
                elif colDataType == "float64":
                    # Add column to compound table
                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" REAL;")
                    # Add default values to the new column
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", (str(0.0), ))
                                                
                    customDataType = str(cdDataTypeDict["Double"])
                    valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                elif colDataType == "bool":
                    # Add column to compound table
                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" NUMERIC;")
                    # Add default values to the new column
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", ("False", ))

                    # Storing bool values as strings in CD is easier
                    customDataType = str(cdDataTypeDict["String"])
                    valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                else:
                    # Add column to compound table
                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" TEXT;")
                    # Add default values to the new column
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", ("", ))
                                             
                    customDataType = str(cdDataTypeDict["String"])
                    valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                
                # We want to make the default column width larger for the Notes column
                if colDisplayName == "Notes":
                    columnWidth = "150"
                else:
                    columnWidth = "-1"
                            
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
                                4, 0, (?), \
                                '', 1, 0);", (compoundTblID, colDBName, customDataType, valueType, colDisplayName, colDisplayName+": This column has been added by CDExcelMessenger.py", columnWidth, ))
                    
                report.append("Column: "+colDisplayName+" added to "+cdResultsFilePath)

            # If the column already exists in CD, make sure the data type matches the Excel column data type
            else:
                if colDisplayName in peakTable.columns:
                    if colDisplayName == "Checked":
                        if colDataType != "bool":
                            raise TypeError("TypeError", "The Checked column in the Excel file must be boolean (True/False)")
                    
                    # We're not checking the Tags column data type here, because that gets checked later
                    elif colDisplayName != "Tags":
                        # Int column
                        if colDataType == "int64":
                            cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                            if cursor.fetchall()[0][0] != cdDataTypeDict["Int64"]:
                                raise TypeError("TypeError", colDisplayName+" data type in "+excelFilePath+" doesn't match data type in "+cdResultsFilePath)

                        # Float column
                        elif colDataType == "float64":
                            cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                            if cursor.fetchall()[0][0] != cdDataTypeDict["Double"]:
                                raise TypeError("TypeError", colDisplayName+" data type in "+excelFilePath+" doesn't match data type in "+cdResultsFilePath)
                
                        # Bool or object column
                        else:
                            cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                            if cursor.fetchall()[0][0] != cdDataTypeDict["String"]:
                                raise TypeError("TypeError", colDisplayName+" data type in "+excelFilePath+" doesn't match data type in "+cdResultsFilePath)

            # Loop through each row in the Excel file, to update the current column in the CD results file
            for row in range(peakRowCount):
                
                # The Excel ID is needed to match rows between the excel file and CD results file
                ID = peakTable.at[row,"compoundID"]
                
                # The Tags column needs to be handled differently
                if colDBName == "Tags":
                    tagString = ""
                    
                    # If there are tags that are only found in the excel 'Tags' column
                    if tagsInTagsCol != []:
                    
                        # If the current Excel cell is not empty
                        if pd.notnull(peakTable.at[row,colDisplayName]):
                            currTagString = peakTable.at[row,colDisplayName]
                            # Make sure the Tags value in Excel is a string
                            if type(currTagString) == str:
                                tagStringSplit = currTagString.split(";")
                                for tag in tagsInTagsCol:
                                    if tag in tagStringSplit:
                                        tagString = tagString + tag + ";"
                            else:
                                raise TypeError("TypeError", str(currTagString)+" in the Tags column should be a string")
                        
                    # If the user wants to update the Tags in CD using values from chosen Excel columns
                    if tagList is not None and tagList != []:
                        
                        # Create a tagString containing Tags that contain True or 1 values in the current Excel row
                        for tag in tagList:
                            if tag in peakTable.columns:
                                tagValue = peakTable.at[row,tag]
                                if tagValue == True or tagValue == 1:
                                    tagString = tagString + tag + ";"                    
                    
                    # Remove the last ;
                    if len(tagString) > 0:
                        tagString = tagString[:-1]
                    
                    # Convert the tag string to bytes that CD can read
                    value = tagStringToBytes(tagString, cdResultsFilePath, cursor)
                    if value == None:
                        value = b"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"

                    # Update the current row and column in the CD results file, also set Cleaned to True    
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?), Cleaned = 'True' WHERE ID = (?);", (value, str(ID), ))  
                    
                    # Update MST column
                    tagStringSplit = tagString.split(";")
                    currMSI = getMSIValue(tagStringSplit)
                    
                    # Update the current row and column in the CD results file
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET MSI = (?) WHERE ID = (?);", (str(currMSI), str(ID), ))  
                    
                                            
                # If the column isn't the Tags column
                else:
                    # The Checked column also needs to be handled differently
                    if colDBName == "Checked": 
                        value = peakTable.at[row,colDisplayName]
                        if bool(value):
                            value = 1
                        else:
                            value = 0
                          
                    # Notes is a default column so we still want to add it even if Notes doesn't exist in the Excel data
                    elif colDBName == "Notes" and "Notes" not in peakTable.columns:
                        value = ""
                    
                    # Not the Tags or Checked column
                    # Also not the Notes column unless the Notes column exists in the Excel data
                    else:
                        value = peakTable.at[row,colDisplayName]
                        
                        # Make sure the numeric values aren't too large
                        if type(value) == int or type(value) == float:
                            if value.bit_length() > 64: 
                                raise ValueError("ValueError", "The value '"+str(value)+"' is too large. 64 bits is the maximum size for numeric values.")
                    
                    # Update the current row and column in the CD results file, also set Cleaned to True
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?), Cleaned = 'True' WHERE ID = (?);", (str(value), str(ID), ))   

            report.append("Column: "+colDisplayName+" updated")
        
        # Save changes to CD database
        conn.commit()
            
        # Close the connection to the Compound Discoverer file
        cursor.close()
        conn.close()
            
        if verbose: 
            report.append(cdResultsFilePath+" updated")
            for i in report:
                print(i)
        else:
            report.append(cdResultsFilePath+" updated")
            return report
    
    # Operational Error 
    except sqlite3.OperationalError:
        # Close the connection to the Compound Discoverer file
        if cursor is not None:
            cursor.close()
        if conn is not None:
            conn.close()
        
        if verbose:
            print("SQLite3:OperationalError: It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
        else:
            raise sqlite3.OperationalError("SQLite3:OperationalError", "It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
      
    # Get info about other errors
    except Exception as e:
        # Close the connection to the Compound Discoverer file
        if cursor is not None:
            cursor.close()
        if conn is not None:
            conn.close()
        
        if verbose:
            print(e)
        else:    
            raise e
    

#####################################################################################
## Function: validateUpdateExcelInput()
#####################################################################################
'''
This function makes sure the user input for updating the Excel results file is all valid

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'excelFilePath' = The path to an Excel file.
'peakSheetName' = The name of the Excel sheet containing the peak data.
'excelColList' = a list of columns in the Excel file that the user wishes to update.
'removeCheckedRows' = Boolean value that controls whether or not rows get deleted in the Excel file if the row is Checked in CD.
'verbose' = Boolean value that controls the output to the console. 
'''

def validateUpdateExcelInput(cdResultsFilePath, excelFilePath, peakSheetName, dataSheetName, excelColList, removeCheckedRows, appendSheets, verbose):
    # Validate 'cdResultsFilePath'
    if type(cdResultsFilePath) != str:
        raise TypeError("TypeError", "Make sure 'cdResultsFilePath' is a string value")
     
    # Validate 'excelFilePath'
    if type(excelFilePath) != str:
        raise TypeError("TypeError", "Make sure 'excelFilePath' is a string value")

    # Validate 'peakSheetName'
    if type(peakSheetName) != str:
        raise TypeError("TypeError", "Make sure 'peakSheetName' is a string value")
        
    # Validate 'dataSheetName'
    if dataSheetName is not None:
        if type(dataSheetName) != str:
            raise TypeError("TypeError", "Make sure 'dataSheetName' is a string value")
    
    # Validate 'excelColList'
    if excelColList is not None:
        if type(excelColList)!= list:
            raise TypeError("TypeError", "Make sure 'excelColList' is a list")
        
        for i in excelColList:
            if type(i) != str:
                raise TypeError("TypeError", "Make sure all values of 'excelColList' are string values")
    
    # Validate 'removeCheckedRows'
    if type(removeCheckedRows) != bool:
        raise TypeError("TypeError", "Make sure 'removeCheckedRows' is a boolean value")
        
    # Validate 'appendSheets'
    if type(appendSheets) != bool:
        raise TypeError("TypeError", "Make sure 'appendSheets' is a boolean value")
    
    # Validate 'verbose'
    if type(verbose) != bool:
        raise TypeError("TypeError", "Make sure 'verbose' is a boolean value")
    
        
#####################################################################################
## Function: updateExcelFile()
#####################################################################################
'''
This function imports data from a Compound Discoverer (CD) results file into an Excel file.

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'excelFilePath' = The path to an Excel file.
'peakSheetName' = The name of the Excel sheet containing the peak table.
'dataSheetName' = The name of the Excel sheet containing the data table.
'excelColList' = A list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file).
'removeCheckedRows' = Boolean value that controls whether or not rows get deleted in the Excel file if the row is Checked in CD.
'appendSheets' = Boolean value that controls whether or not the Excel data produced by this function gets appended to the Excel file,
    True = append sheets, False = overwrite sheets
'verbose' = Boolean value that controls the output to the console. 
    If False, hide outputs and return the outputs as a list.
    
OUTPUT:
'report' = A list of messages that can be printed to console if 'verbose' is True.
'''
    
def updateExcelFile(cdResultsFilePath, excelFilePath, peakSheetName, dataSheetName = None, excelColList = None, removeCheckedRows = False, appendSheets = False, verbose = True):
    # Set the sqlite connection and cursor variables to None so that they can be closed during exceptions
    # only if the exception happened after the connections have been set
    conn = None
    cursor = None
    try:
        # Basic validation on user input
        validateUpdateExcelInput(cdResultsFilePath, excelFilePath, peakSheetName, dataSheetName, excelColList, removeCheckedRows, appendSheets, verbose)
        
        report = []
    
        if verbose == True: 
            print("\nImporting data from "+cdResultsFilePath+" into "+excelFilePath)
    
        # If the results file can't be found
        if os.path.exists(cdResultsFilePath) == False:
            raise FileNotFoundError("FileNotFoundError", cdResultsFilePath+" can't be found")

        # Open connection to the Compound Discoverer File
        conn = sqlite3.connect(cdResultsFilePath)
        cursor = conn.cursor()
        
        # Get Peak table
        try:
            # Get Excel data in a dataframe, fill NA values in that dataframe, and get the number of rows in that dataframe
            peakTable = pd.read_excel(excelFilePath, sheet_name = peakSheetName)            
            peakTable = fillNAValuesInDF(peakTable)
            peakRowCount = len(peakTable.index)
        
        # If the Excel file doesn't have the correct sheet
        except ValueError:
            raise ValueError("ValueError", "Can't find "+peakSheetName+" in "+excelFilePath)
        
        # If the Excel file can't be found
        except FileNotFoundError:
            raise FileNotFoundError("FileNotFoundError", "Can't find "+excelFilePath)
    
        # If permission to the Excel file was denied
        except PermissionError:
            raise PermissionError("PermissionError", "Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")

        # Get Data table
        if dataSheetName is not None:
            try:
                # Get data sheet
                dataTable = pd.read_excel(excelFilePath, sheet_name = dataSheetName)            
        
            # If the Excel file doesn't have the correct sheet
            except ValueError:
                raise ValueError("ValueError", "Can't find "+dataSheetName+" in "+excelFilePath)
        
            # If the Excel file can't be found
            except FileNotFoundError:
                raise FileNotFoundError("FileNotFoundError", "Can't find "+excelFilePath)
    
            # If permission to the Excel file was denied
            except PermissionError:
                raise PermissionError("PermissionError", "Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")
        
        # If the MSI column isn't in the peak sheet, add it
        if "MSI" not in peakTable.columns:
            peakTable["MSI"] = 0
            report.append("Column: MSI added to "+excelFilePath)
        
        # If the Excel data doesn't contain the CD database IDs, add them
        if "compoundID" not in peakTable.columns:
            peakTable, newReport = createCompoundIDColumns(cdResultsFilePath, conn, cursor, peakRowCount, excelFilePath, peakTable, peakSheetName)
            for i in newReport:
                report.append(i)
        
        # If the user chose to import Tag data into Excel columns
        if "Tags" in excelColList:
            # Get the IDs of visible tags from CD results file 
            cursor.execute("SELECT BoxID FROM DataDistributionBoxExtendedData WHERE name = 'EntityItemTagVisibility' AND ValueString = 'True';")
            visibleIDList = []
            for ID in cursor.fetchall():
                visibleIDList.append(ID[0])
                 
            # 'keys' will be used in an SQLite statement to get the names of visible tags     
            keys = ""
            # Loop through the list of visible Tag IDs
            for ID in visibleIDList:
                # Add the Tag ID to keys
                keys += "'"+str(ID)+"', "
            
            # Remove last ', ' from the keys
            keys = keys[:-2] 
                 
            # Get the names of visible tags     
            visibleTagList = []
            cursor.execute("SELECT Name FROM DataDistributionBoxes WHERE BoxID IN ("+keys+");")
            for tag in cursor.fetchall():
                visibleTagList.append(tag[0])

            # 'tagList' will hold the names of Tags
            # 'defaultValList' will be a list of False values,
            # 'newColList' will contain the new columns to be added to the Excel file
            tagList = []
            defaultValList = [False] * peakRowCount
            newColList = []
            
            # Get binary columns from the peak Table
            binaryCols = peakTable.columns[peakTable.isin([0,1]).all()]
            # Loop through the visible tags
            for tag in visibleTagList:
                
                # If the current tag is not in the Excel file, add it
                if tag not in peakTable.columns:
                    newColList.append(pd.DataFrame (defaultValList, columns = [tag]))
                    tagList.append(tag)
                    report.append("Column: "+tag+" added to "+excelFilePath)
                
                # If the current tag is in the Excel file
                else:
                    # If the Excel column is boolean
                    if peakTable.dtypes[tag] == "bool":
                        tagList.append(tag)
                        report.append("Column: "+tag+" updated in "+excelFilePath)
                    
                    # If the Excel column is binary
                    elif peakTable.dtypes[tag] == "int64":
                        if tag in binaryCols:    
                            tagList.append(tag)
                            report.append("Column: "+tag+" updated in "+excelFilePath)
                        else:
                            report.append("WARNING: column - "+tag+" already exists in "+excelFilePath+", but is not a boolean or binary column")
                        
                    # If the Excel column is not bool or binary
                    else:
                        report.append("WARNING: column - "+tag+" already exists in "+excelFilePath+", but is not a boolean or binary column")
            
            # If there was at least one visible tag
            if newColList != []:
                # Add Excel data to 'newColList', so that we can concatinate all the columns together 
                newColList.append(peakTable)
                peakTable = pd.concat(newColList, axis=1, ignore_index=False)
            
        # Get list of tuples
        # Each tuple will contains the column DB name and display name
        colNameTupleList, newReport = getColNamesForUpdatingExcel(cdResultsFilePath, cursor, peakTable, excelColList)
        for i in newReport:
            report.append(i)

        # Get the ID of compound table
        cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
        compoundTblID = cursor.fetchall()[0][0]
 
        # Loop through each column tuple in the list of tuples, to update each column in the list
        for colNameTuple in colNameTupleList:
                
            # Get the column DB name and display name
            colDBName = colNameTuple[0]
            colDisplayName = colNameTuple[1]
            
            # If the column is not in the Excel file, add the column with placeholder values
            if colDisplayName not in peakTable.columns:
                peakTable[colDisplayName] = [""]*peakRowCount
                
                report.append("Column: "+colDisplayName+" added to "+excelFilePath)
                
            # If the column is already in the Excel file
            else:
                report.append("Column: "+colDisplayName+" updated in "+excelFilePath)
                     
            # Check if the column data type is boolean, the colIsBool variable gets used to make sure the column stays boolean
            colIsBool = False
            if peakTable.dtypes[colDisplayName] == "bool":
                colIsBool = True
                
            # Loop through each row in the excel file
            for row in range(peakRowCount):
                   
                # The Excel ID is needed to match rows between the excel file and CD results file
                ID = peakTable.at[row,"compoundID"]
                        
                # Get the value of the current row and column from the CD results file    
                cursor.execute("SELECT "+colDBName+" FROM ConsolidatedUnknownCompoundItems WHERE ID = (?);", (str(ID), ))     
                value = cursor.fetchall()[0][0]
                
                # The Tags column needs to be handled differently
                if colDBName == "Tags":
                   
                    # Update the current row and column in the Excel data frame after converting the bytes value to a string
                    tagString = tagBytesToString(value, cdResultsFilePath, cursor)
                    peakTable.at[row,colDisplayName] = tagString                        
                    
                    # Create a list of tags, ';' is the delimiter
                    tagStringSplit = tagString.split(";")
                    
                    # Update the individual Tag columns
                    if tagList != []:                
                        # Loop through the list of all tags
                        for tag in tagList:
                            
                            # If the current tag is in 'tagStringSplit', 
                            # that means the current tag for this row is checked in the CD results file,
                            # and we can update the tag column in the Excel file
                            if tag in tagStringSplit:
                                # update the tag column in the Excel file
                                peakTable.at[row,tag] = True
                            else:
                                peakTable.at[row,tag] = False
                                
                    # Update MSI column
                    if peakTable.dtypes["MSI"] == "int64":
                        currMSI = getMSIValue(tagStringSplit)
                        peakTable.at[row,"MSI"] = currMSI

                # If the column is not the Tags column
                else:
                    # if the current column is a boolean column
                    if colIsBool:
                        
                        # If the value is stored as a string in CD
                        if type(value) == str:
                            if value.upper() == "TRUE":
                                value = True
                                    
                            else:
                                value = False
                            
                        # If the value is not stored as a string in CD
                        else:
                            value = bool(value)
                    
                        # Update the current row and column in the Excel data
                        peakTable.at[row,colDisplayName]=value
                    
                    # If the column is not a boolean column
                    else:
                        # Make sure the numeric values aren't too large
                        if type(value) == int or type(value) == float:
                            if value.bit_length() >= 64: 
                                raise ValueError("ValueError", "The value '"+str(value)+"' is too large. 64 bits is the maximum size for numeric values.")

                        # Update the current row and column in the Excel data
                        peakTable.at[row,colDisplayName]=value

            # This is to make sure boolean columns are set as bool in the Excel file
            if colIsBool:                 
                peakTable[colDisplayName] = peakTable[colDisplayName].astype('bool')
        
        # If the user wishes to drop rows that have been checked
        if removeCheckedRows:
            if "Checked" in peakTable.columns:
                if peakTable.dtypes["Checked"] == "bool":
                    uidList = []
                    # If the user provided the data sheet, we need to get the 'UID' of the rows we are going to drop from the peak sheet
                    if dataSheetName is not None:
                        if "UID" in peakTable.columns:
                            if peakTable.dtypes["UID"] == "object":
                                # Loop through the rows of the peak Table to get the 'UID' values of rows that are being dropped
                                for row in range(peakRowCount):
                                    if peakTable.at[row, "Checked"] == True:
                                        currUID = peakTable.at[row, "UID"]
                                        if currUID in dataTable.columns:
                                            uidList.append(currUID) 
                    
                    # Drop rows from the peak table
                    peakTable.drop(peakTable[peakTable.Checked == True].index, inplace=True)
                    report.append("Dropped rows from peak sheet")
            
                    if uidList != []:
                        # Drop columns from data sheet
                        dataTable.drop(uidList, axis=1, inplace=True)
                        report.append("Dropped columns from data sheet")
            
        
        # Order Excel columns, put tag columns after the 'Tags' column
        firstCols = []
        if "Idx" in peakTable.columns:
            firstCols.append("Idx")
        if "compoundID" in peakTable.columns:
            firstCols.append("compoundID")
        if "UID" in peakTable.columns:
            firstCols.append("UID")
        if "CIMCBlib" in peakTable.columns:
            firstCols.append("CIMCBlib")
        if "Name" in peakTable.columns:
            firstCols.append("Name")
        if "Notes" in peakTable.columns:
            firstCols.append("Notes")
        if "MSI" in peakTable.columns:
            firstCols.append("MSI")
        if "Tags" in peakTable.columns:
            firstCols.append("Tags")
        for i in tagList:
            firstCols.append(i)
                
        peakTable = peakTable[firstCols + [c for c in peakTable if c not in firstCols]] 
        
        if appendSheets:
            peakSheetName = peakSheetName + "Appended"
            if dataSheetName is not None:
                dataSheetName = dataSheetName + "Appended"
        
        try:
            # Update the Excel file 
            with pd.ExcelWriter(
                excelFilePath,
                mode="a",
                engine="openpyxl",
                if_sheet_exists="replace",
            ) as writer:
                peakTable.to_excel(writer, sheet_name=peakSheetName, index=False) 
            
            if dataSheetName is not None:
                try:
                    # Update the Excel file 
                    with pd.ExcelWriter(
                        excelFilePath,
                        mode="a",
                        engine="openpyxl",
                        if_sheet_exists="replace",
                    ) as writer:
                        dataTable.to_excel(writer, sheet_name=dataSheetName, index=False) 
        
                # If the Excel file doesn't have the correct sheet
                except ValueError:
                    raise ValueError("ValueError", "Can't find "+dataSheetName+" in "+excelFilePath)
        
                # If the Excel file can't be found
                except FileNotFoundError:
                    raise FileNotFoundError("FileNotFoundError", "Can't find "+excelFilePath)
    
                # If permission to the Excel file was denied
                except PermissionError:
                    raise PermissionError("PermissionError", "Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")  

        
        # If the Excel file doesn't have the correct sheet
        except ValueError:
            raise ValueError("ValueError", "Can't find "+peakSheetName+" in "+excelFilePath)
        
        # If the Excel file can't be found
        except FileNotFoundError:
            raise FileNotFoundError("FileNotFoundError", "Can't find "+excelFilePath)
    
        # If permission to the Excel file was denied
        except PermissionError:
            raise PermissionError("PermissionError", "Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")  

        # Close the connection to the Compound Discoverer file
        cursor.close()
        conn.close()
        
        if verbose: 
            report.append(excelFilePath+" updated")
            for i in report:
                print(i)
        else:
            report.append(excelFilePath+" updated")
            return report
        
    # Operational Error 
    except sqlite3.OperationalError:
        # Close the connection to the Compound Discoverer file
        if cursor is not None:
            cursor.close()
        if conn is not None:
            conn.close()
        
        if verbose:
            print("SQLite3:OperationalError: It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
        else:
            raise sqlite3.OperationalError("SQLite3:OperationalError", "It's possible the connection to "+cdResultsFilePath+" was interrupted, or multiple processed are trying to access the same file, or the CD results file has been corrupted.")
    
    # Get info about other errors
    except Exception as e:
        # Close the connection to the Compound Discoverer file
        if cursor is not None:
            cursor.close()
        if conn is not None:
            conn.close()
        
        if verbose:
            print(e)
        else:    
            raise e  
 
 
#####################################################################################
## CleanupPeakTable()
#####################################################################################
'''
This function cleans the peak table and produces stats

INPUT:
'peakTable' = The peak table.
'colsToKeepDict' = Dictionary containing peak columns to keep and the columns new name.
'optionsDict' = Dictionary containing options and the users chosen values for those options.

RETURN:
'peakTable' = Peak table after cleaning.
'report' = Output to print to console is Verbose is True.
'stats' = Stast produced by this function.
'''

def CleanupPeakTable(peakTable,colsToKeepDict,optionsDict):
    
    newDict = {}
    report = [] 

    # Loop through each entry in the colsToKeepDict
    for name, newName in colsToKeepDict.items():

        # If the user has indicated that the current entry is associated with multiple columns in the Peak table
        if name.startswith("<") and name.endswith(">"):
            name = name[1:-1]

            # Get a list of indexes for the columns that start with 'name'
            found = []
            for i in range(len(peakTable.columns)):
                if peakTable.columns[i].startswith(name):
                    found.append(i)

            # If no columns in the Peak table started with 'name'
            if found == []:
                report.append("WARNING: "+name+" wasn't found in the Peak table")

            # If at least one column in the Peak table starts with 'name'
            else:

                # For each column in the Peak table that starts with 'name',
                # add a new entry in 'NewDict' with the original column name as the key,
                # and a reformated column name as the value of that entry
                for j in range(len(found)):
                    label = Peak.columns[found[j]]
                    label = label.replace(name, "", 1)
                    label = label.strip()      
                    newDict[peakTable.columns[found[j]]] = newName+label

        # If the current entry is supposed to be associated with just one column in the Peak table    
        else:
            # Get a list of indexes for the columns that have the name 'name'
            found = []
            for i in range(len(peakTable.columns)):
                if peakTable.columns[i] == name:
                    found.append(i)

            # If no columns in the Peak table have the name 'name'
            if found == []:
                report.append("WARNING: "+name+" wasn't found in the Peak table")

            # If multiple columns in the Peak table have the name 'name'
            elif len(found) > 1:
                report.append("ERROR multiple entries of: "+name)

            # If exactly one column in the Peak table have the name 'name'
            else:
                newDict[peakTable.columns[found[0]]] = newName
    
    # Get the correct columns in a list
    cols = ['Idx', 'UID', 'Name']
    for key in newDict:
        if key not in cols:
            cols.append(key)
            
    # Make sure Peak only has the columns we need
    peakTable = peakTable[cols]
    peakTable = peakTable.rename(columns=newDict)
    
    # Create CIMCBlib column and move the CIMCB ID code.
    if optionsDict["CIMCBlib"]:

        # Tracks if the CIMCBlib column is in Peak
        CIMCBlibInPeak = False

        # Loop through each row of Peak and get the current Label.     
        for i in range(len(peakTable.index)):
            currLabel = peakTable.at[i,"Name"]
            
            if type(currLabel) == str:
            
                # If the current label starts with 'ECU', add the CIMCBlib column if it doesn't exist,
                if currLabel.startswith("ECU"):
                    if CIMCBlibInPeak == False:
                        if 'CIMCBlib' not in peakTable.columns:
                            # Add new column to Peak with empty values
                            peakTable["CIMCBlib"] = ""
                            CIMCBlibInPeak = True
                        else:
                            CIMCBlibInPeak = True
                        
            
                    # Split the current label and update both 'CIMCBlib' and 'Label' in the current row of Peak
                    temp = currLabel.split("_", 1)
                    peakTable.at[i,"CIMCBlib"] = temp[0]
                    peakTable.at[i,"Name"] = temp[1]    
        
        if 'CIMCBlib' in peakTable.columns:
            # Make sure the first four columns of Peak are 'Idx', 'UID', 'CIMCBlib', and 'Name'
            peakTable = peakTable[['Idx', 'UID', 'CIMCBlib', 'Name'] + [c for c in peakTable if c not in ['Idx', 'UID', 'CIMCBlib', 'Name']]] 
        else:
            # Make sure the first three columns of Peak are 'Idx', 'Name', and 'Label'
            peakTable = peakTable[['Idx', 'UID', 'Name'] + [c for c in peakTable if c not in ['Idx', 'UID', 'Name']]] 
            
    stats = [];
            
    # Get the sum of rows that have a MS2 hit
    if optionsDict["MSHit"]:
    
        # Add new column to Peak table with False values, 
        # these values will be changed to True during the upcoming loop if there is a ms2Hit
        peakTable["ms2Hit"] = False
        
        # If the MS2 column is in Peak, loop through each row of Peak, and track ms2Hits
        if 'MS2' in peakTable.columns:
            ms2hits = 0
            for i in range(len(peakTable.index)):
                if peakTable.at[i,"MS2"] != "No MS2":
                    ms2hits = ms2hits + 1
                    peakTable.at[i,"ms2Hit"] = True
            
            stats.append(str(ms2hits) + " peaks with MS2 spectra")

    # Peak cleaning stats.

    # Get the indexes for the column that start with 'mzList_'
    found = -1
    for i in range(len(peakTable.columns)):
        if peakTable.columns[i].startswith("mzList_"):
            found = i
            
    # Get the number of mass hits
    massHit = []
    massHitSum = 0
    if found != -1:
        for i in range(len(peakTable.index)):
            if peakTable.iloc[i, found] != "No matches found":
                massHit.append(1)
                massHitSum = massHitSum + 1
            else:
                massHit.append(0)
    else:
        massHit = [0] * len(peakTable.index)

    # Get the number of vault hits
    vaultHit = []
    vaultHitSum = 0
    if "mzVaultMatch" in peakTable.columns:   
        for i in range(len(peakTable.index)):
            if float(peakTable.at[i, "mzVaultMatch"]) > optionsDict["mzmatch"]:
                vaultHit.append(1)
                vaultHitSum = vaultHitSum + 1
            else:
                vaultHit.append(0)
    else:
        vaultHit = [0] * len(peakTable.index)
        
    # Get the number of cloud hits
    cloudHit = []
    cloudHitSum = 0
    if "mzCloudMatch" in peakTable.columns:   
        for i in range(len(peakTable.index)):
            if float(peakTable.at[i, "mzCloudMatch"]) > optionsDict["mzmatch"]:
                cloudHit.append(1)
                cloudHitSum = cloudHitSum + 1
            else:
                cloudHit.append(0)
    else:
        cloudHit = [0] * len(peakTable.index)
            
    # Get the number of mass, vault, and cloud hits        
    peakTable["Hit"] = 0
    hitSum = 0
    for i in range(len(peakTable.index)):
        hit = cloudHit[i] + vaultHit[i]*10 + 100*massHit[i] 
        peakTable.at[i, "Hit"] = hit
        if hit > 0:
            hitSum = hitSum + 1

    stats.append(str(massHitSum) + " MassList hits")
    stats.append(str(vaultHitSum) + " mzVault hits")
    stats.append(str(cloudHitSum) + " mzCloud hits")
    if hitSum == 1:
        stats.append(str(hitSum) + " unique hit")
    else:
        stats.append(str(hitSum) + " unique hits")

    return peakTable, report, stats


#####################################################################################
## MergeMetaintoData()
#####################################################################################
'''
This function merges Meta columns into the Data table

INPUT:
'dataTable' = The Data table.
'metaTable' = The Meta table.

OUTPUT:
'dataTable' = The Data table after Meta columns have been added.
'''

def MergeMetaintoData(dataTable, metaTable):

    # Make sure Meta table contains the necessary columns
    metaHeader = metaTable.columns
    if all(item in metaHeader for item in ["Filename", "Batch", "Order", "SampleID", "SampleType"]) == False:
        raise Exception("MetaTable must contain columns: 'Filename','Batch','Order','SampleID' & 'SampleType'")
    
    # Make sure Data table contains the necessary columns    
    dataHeader = dataTable.columns            
    if all(item in dataHeader for item in ["Idx", "Filename"]) == False:
        raise Exception("DataTable must contain columns: 'Idx','Filename'")        
    
    # Sort Data and Meta tables by the 'Filename' column, 
    # and check that the 'Filename' columns are identical in both tables
    tempData = dataTable.sort_values(by=["Filename"])
    tempMeta = metaTable.sort_values(by=["Filename"])
    dataFilenames = tempData["Filename"]        
    metaFilenames = tempMeta["Filename"]
    if dataFilenames.equals(metaFilenames) == False:
        raise Exception("MetaTable & DataTable must have identical Filenames")
    
    # Drop 'Filename' from the Meta table becasue we don't need it any more,
    # but we do still need it in the Data table
    tempMeta = tempMeta.drop(columns=['Filename'])
    
    # If their are any duplicate columns still in the Data and Meta tables,
    # remove that column from the Data table
    for i in tempData.columns:
        if i in tempMeta.columns:
            tempData = tempData.drop(columns=i)
        
    # Concatinate the Data and Meta tables,
    # there should be no duplicate columns now
    tempData = pd.concat([tempData, tempMeta], axis=1)
    
    # Make sure the first two columns of Data are 'Idx', and 'Filename',
    # then the next columns should be the Meta columns, followed by the other Data columns
    tempCols = ['Idx', 'Filename'] 
    for i in tempMeta.columns:
        tempCols.append(i)
    tempData = tempData[tempCols + [c for c in tempData if c not in tempCols]] 
    
    # Sort Data by 'Idx' before returning Data
    dataTable = TempData.sort_values(by=['Idx'])
    return dataTable


#####################################################################################
## validatingDataPeakTables()
#####################################################################################
'''
This function validates the Data table and Peak table.

INPUT:
'dataTable' = The Data table.
'peakTable' = The Peak table.
'optionsDict' = Dictionary containing options and the values the user chose for those options.

OUTPUT:
'dataTable' = The Data table after validation.
'peakTable' = The Peak table after validation.

'''

def validatingDataPeakTables(dataTable, peakTable, optionsDict):

    #column names
    peakHeader = peakTable.columns

    # if UID or Name not in peak columns
    if "UID" not in peakHeader or "Name" not in peakHeader:
        raise Exception("TidyData:PeakTableError", "PeakTable must contain columns ''UID'' & ''Name''")

    # Make sure all values in the 'UID' column are unique
    peakList = peakTable["UID"]
    peaks = peakList.duplicated()
    if True in peaks.values:
        raise Exception("TidyData:PeakTableError", "All ''Names'' in the PeakSheet must be unique.")
    
    # Make sure the peak names are identical in the DataSheet and PeakSheet
    peakList = peakList.values
    dataHeader = dataTable.columns
    dataHeader = dataHeader.values
    temp = []
    for i in dataHeader:
        if i in peakList:
            temp.append(i)
    if len(peakList) != len(temp):
        raise Exception("TidyData:DataTableError", "The peak names in the DataSheet should be unique, and exactly match the peak names in the PeakSheet.")
    else:
        for i in range(len(peakList)):
            if peakList[i] != temp[i]:
                raise Exception("TidyData:DataTableError", "The peak names in the DataSheet should be unique, and exactly match the peak names in the PeakSheet.")

    # Make sure SampleID, SampleType, Order, and Batch are in data columns
    if "SampleID" not in dataHeader or "SampleType" not in dataHeader or "Order" not in dataHeader or "Batch" not in dataHeader:
        raise Exception("TidyData:DataTableError", "DataTable must contain columns ''SampleID'', ''SampleType'', ''Order'', & ''Batch''")

    # Make sure SampleIDs are strings
    if dataTable.dtypes["SampleID"] != "object":
        dataTable["SampleID"] = dataTable["SampleID"].apply(str)
    
    # A list containing the valid sample types
    validSampleTypes = ["Sample", "QC", "Blank", "Reference"]
    
    # Make sure the Data table is sorted by Batch then Order,
    # and make sure the Order values are increasing integers
    dataTable = dataTable.sort_values(by=['Batch', 'Order'])
    prevOrder = -1
    prevBatch = -1
    batchSum = 0
    for i in range(len(dataTable.index)):
        currOrder = dataTable.at[i, "Order"]
        if currOrder.dtype == "int64" and currOrder > prevOrder:
            prevOrder = currOrder
        else:
            raise Exception("TidyData:QCRSCDataTableError", "DataTable ''Order'' column must contain unique increasing integer values matching to increasing batch number.")
        
        # Make sure the Batch values are integers with no missing values
        currBatch = dataTable.at[i, "Batch"]
        if currBatch.dtype != "int64" or currBatch is None:
            raise Exception("QCRSC:DataTableError", "DataTable ''Batch'' column must contain integer values and no missing values")
        if currBatch > prevBatch:
            batchSum = batchSum + 1
            prevBatch = currBatch
            
        # Make sure the Sample types are valid
        currSampleType = dataTable.at[i, "SampleType"]
        if currSampleType not in validSampleTypes:
            raise Exception("QCRSC:QCRSCDataTableError", "DataTable ''SampleType'' column values must be one of the following: ''Sample'',''QC'',''Blank'', or ''Reference''")
        
    # Remove Data columns that share names with sample types
    if "QC" in dataHeader:
        dataTable = dataTable.drop(columns=['QC'])
    if "Reference" in dataHeader:
        dataTable = dataTable.drop(columns=['Reference'])
    if "Blank" in dataHeader:
        dataTable = dataTable.drop(columns=['Blank'])
    if "Sample" in dataHeader:
        dataTable = dataTable.drop(columns=['Sample'])
        
    # Add sample types columns and set to boolean values 
    qcSum = 0
    for i in range(len(dataTable.index)):
        currSampleType = dataTable.at[i, "SampleType"]
        if currSampleType == "QC":    
            qcSum = qcSum + 1
            dataTable.loc[i,["QC","Blank","Reference","Sample"]] = [True,False,False,False]         
        elif currSampleType == "Blank":
            dataTable.loc[i,["QC","Blank","Reference","Sample"]] = [False,True,False,False]
        elif currSampleType == "Reference":
            dataTable.loc[i,["QC","Blank","Reference","Sample"]] = [False,False,True,False]
        elif currSampleType == "Sample":
            dataTable.loc[i,["QC","Blank","Reference","Sample"]] = [False,False,False,True]

    # Make sure their are at least 3 QCs samples for each Batch
    if qcSum < batchSum*3:
        raise Exception("QCRSC:DataTableError", "There has to be 3 or more QCs per batch for this data to be valid for QC assessment!")
    
    # Make sure the column in the Data table are ordered correctly,
    # the order will be 'firstCols', 'midCols', then the 'peakList'
    midCols = ["SampleType", "Order", "Batch", "QC", "Blank", "Reference", "Sample"]
    firstCols = []
    for i in dataTable.columns:
        if i not in peakList and i not in midCols:
            firstCols.append(i)
    firstCols = firstCols + midCols
    dataTable = dataTable[firstCols + [c for c in dataTable if c not in firstCols]] 
    
    return dataTable, peakTable


#####################################################################################
## validatingTidyDataInput()
#####################################################################################
'''
This function makes sure the user input for creating the TidyData is all valid

INPUT:
'excelFilePath' = The path to an Excel file.
'colsToKeepDict' = Dictionary containing the Peak columns the user wants to keep and the name the column should be renamed to.
'optionsDict' = Dictionary containing options and the user choices to those options.
'verbose' = Boolean value that controls the output to the console. 
'''

def validateTidyDataInput(excelFilePath, colsToKeepDict, optionsDict, verbose):
    # Validate 'excelFilePath'
    if type(excelFilePath) != str:
        raise TypeError("TypeError", "Make sure 'excelFilePath' is a string value")
    
    # Validate 'colsToKeepDict'
    try:
        for key, value in colsToKeepDict.items():
            if type(key) != str:
                raise ValueError("ValueError", "Make sure all keys in 'colsToKeepDict' are string values")
            if type(value) != str:
                raise ValueError("ValueError", "Make sure all values in 'colsToKeepDict' are string values")
                
    # If 'colsToKeepDict' isn't a dictionary
    except AttributeError:
        raise ValueError("ValueError", "Make sure 'colsToKeepDict' is a dictionary")
    
    # Validate 'optionsDict'
    try:
        if type(optionsDict["CIMCBlib"]) != bool:
            raise TypeError("TypeError", "Make sure you set 'CIMCBlib' to a boolean value")
        if type(optionsDict["MSHit"]) != bool:
            raise TypeError("TypeError", "Make sure you set 'MSHit' to a boolean value")
        if type(optionsDict["mzmatch"]) != int and type(optionsDict["mzmatch"] != float):
            raise TypeError("TypeError", "Make sure you set 'mzmatch' to a float or an integer value")
        if type(optionsDict["UIDPrefix"]) != str:
            raise TypeError("TypeError", "Make sure you set 'UIDPrefix' to a string value")

    # If 'optionsDict' isn't a dictionary
    except TypeError:
        raise TypeError("TypeError", "Make sure 'optionsDict' is a dictionary")
        
    # If one of the Keys can't be found
    except KeyError:
        raise KeyError("KeyError", "Make sure 'optionsDict' has the keys 'CIMCBlib', 'MSHit', 'mzmatch', 'ColumnWarning', and 'UIDPrefix'")

    # Validate 'verbose'
    if type(verbose) != bool:
        raise TypeError("TypeError", "Make sure 'verbose' is a boolean value")


#####################################################################################
## tidyData()
#####################################################################################
'''
This function converts an Excel file into the TidyData format.
This function expects the Excel file to have a Compounds table produced by CD and 
a Meta table created from that Compounds table.

INPUT:
'excelFilePath' = The path to an Excel file.
'colsToKeepDict' = Dictionary containing the Peak columns the user wants to keep and the name the column should be renamed to.
'optionsDict' = Dictionary containing options and the user choices to those options.
'verbose' = Boolean value that controls the output to the console. 

OUTPUT:
'report' = A list of messages that can be printed to console if 'verbose' is True.
'stats' = A list of stats produced by the CleanupPeakTable function.
'''

def tidyData(excelFilePath, colsToKeepDict, optionsDict, verbose = True):
    try:
        validateTidyDataInput(excelFilePath, colsToKeepDict, optionsDict, verbose)
        
        try:
            # Get data from Excel file
            compTable = pd.read_excel(excelFilePath, sheet_name = "Compounds")
            peakTable = pd.read_excel(excelFilePath, sheet_name = "Compounds")
            metaTable = pd.read_excel(excelFilePath, sheet_name = "Meta")
        
        # If the Excel file doesn't have the correct sheets
        except ValueError:
            raise ValueError("ValueError", "Make sure "+excelFilePath+" has the Compounds and Meta sheets")
       
        # If the Excel file can't be found
        except FileNotFoundError:
            raise FileNotFoundError("FileNotFoundError", "Can't find "+excelFilePath)
    
        # If permission to the Excel file was denied
        except PermissionError:
            raise PermissionError("PermissionError", "Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")
                
        # Make sure the original name is in the peak table
        if "Name" not in peakTable.columns:
            raise ValueError("ValueError", "The original Name column can't be found in "+excelFilePath)
        
        prefix = optionsDict["UIDPrefix"]
        
        # Create lists containing peak Idx and the UID data
        peakIdxList = list(range(1, len(peakTable.index) + 1))
        uidList = [prefix + str(x) for x in peakIdxList]
        
        # Create new columns called 'Idx' and 'UID'  
        peakTable["Idx"] = peakIdxList
        peakTable["UID"] = uidList        
        
        # Create Data dataframe with Filename and Idx data
        dataTable = pd.DataFrame(data=metaTable["Filename"])
        dataIdxList = list(range(1, len(dataTable.index) + 1))
        dataTable["Idx"] = dataIdxList
        
        areaDict = {}

        index = 0
        # Loop through the peak table and add the Area data to a dictionary
        for col in peakTable.columns:
            if col.startswith("Area: "):
                areaDict[index] = peakTable.loc[:, col]
                index = index + 1
    
        if areaDict == {}:
            raise ValueError("ValueError", "The Area columns can't be found in "+excelFilePath)
    
        # Create Area data frame, tranpose, change column names, and add to the Data data frame
        areaTable = pd.DataFrame(data=areaDict)
        areaTable = areaTable.T
        areaTable.columns = uidList
        dataTable = pd.concat([dataTable, areaTable], axis=1)
    
        report = []
        stats = []
    
        # Merge Meta table data into the Data table
        dataTable = MergeMetaintoData(dataTable, metaTable)
        
        # Clean Peak table
        peakTable, report, stats = CleanupPeakTable(peakTable,colsToKeepDict,optionsDict)
        
        # Validate the Data table and Peak table
        dataTable, peakTable = validatingDataPeakTables(dataTable, peakTable, optionsDict)
        
        try:
            # Append Meta sheet, Data sheet, and Peak sheet to excel file
            with pd.ExcelWriter(excelFilePath, engine="openpyxl") as writer:  
                compTable.to_excel(writer, sheet_name='Compounds', index=False)
                metaTable.to_excel(writer, sheet_name='Meta', index=False)
                dataTable.to_excel(writer, sheet_name='Data', index=False)
                peakTable.to_excel(writer, sheet_name='Peak', index=False)
       
        # If the Excel file can't be found
        except FileNotFoundError:
            raise FileNotFoundError("FileNotFoundError", "Can't find "+excelFilePath)
    
        # If permission to the Excel file was denied
        except PermissionError:
            raise PermissionError("PermissionError", "Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")
        
        report.append(excelFilePath+" updated")
        if verbose:
            for i in report:
                print(i)
            print("\nStats:")
            for i in stats:
                print(i)
        else:
            return report, stats
    
    # Print error messages to console if varbose is true,
    # otherwise raise an exception
    except Exception as e:
        if verbose:
            print(e)
        else:    
            raise e
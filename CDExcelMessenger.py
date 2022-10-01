#####################################################################################
## Import Modules
#####################################################################################

import sqlite3
import os.path
import re
import logging
logger = logging.Logger('catch_all')
import pandas as pd
    
    
#####################################################################################
## Function: binarySearch()
#####################################################################################
'''   
A binary search function to improve performance when searching an ordered list.

INPUT:
'value' = The value that this function is searching for.
'searchList' = An ordered list.

OUTPUT:
'mid' = The index in the search list that the value is found.
    Will return -1 if the value is not found in the search list.
'''

def binarySearch(value, searchList):
    low = 0
    high = len(searchList) - 1
    
    # While there is still search space left in the search list
    while low <= high:

        # Get the middle index of the list
        mid = (high + low) //2
        
        # If the value is found
        if searchList[mid] == value:
            return mid

        # If the value is on the right side of the current search space
        elif searchList[mid] < value:
            low = mid + 1

        # If the value is on the left side of the current search space
        else:
            high = mid - 1

    # The value is not in the search list
    return -1   
   

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
## Function: changeTagNamesAndVisability()
#####################################################################################
'''
This function changes the name of tags in the Compound Discoverer (CD) results file 
based on the names of columns in the Excel file that the user has chosen.

This functon alse changes the Tags visability. If the user has chosen the number of visable Tags,
then the Tag visability is basen on that number. Otherwise, the Tag visability is based on 
the number of Tags in the Tag dictionary.

This function also returns a dictionary of user chosen Tags and thresholds if the
Excel file has a column with that name and if that column is the right type (boolean, integer, or float).

INPUT:
'cursor' = An SQLite cursor.
'excelData' = A DataFrame containing Excel data.
'excelFilePath' = The path to an Excel file.
'tagDict' = A dictionary. The keys of the dictionary are columns in the Excel file that the user wishes to use as Tags in CD,
    the values in the dictionary are thresholds that the user has chosen for checking the Tag boxes.
'tagVisability' = An integer from 0-15 setting the number Tags that the user wants to be visable in CD.
    Will contain a None value if the user hasn't chosen a number.
'verbose' = Boolean value that controls the output to the console. 
    If False, hide outputs and return the outputs as a list.
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
    
OUTPUT:
'tagDictOutput' = Dictionary containing the keys and values from 'tagDict', but only for the chosen Tags that can be used as Tags in CD.
    Will be None if none of the chosen tags can be tags in CD.
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
'''

def changeTagNamesAndVisability(cursor, excelData, excelFilePath, tagDict, tagVisability, verbose, outputList):
    try:
        # If the user has chosen a number of Tags to be visable
        if tagVisability is not None:
        
            # If 'tagVisability' is the wrong type or value
            if type(tagVisability) != int or tagVisability < 0 or tagVisability > 15:
            
                # Set 'tagVisability' to None so that tag visability can be based on the number of user chosen tags that can be tags in CD
                tagVisability = None
                
                if verbose:
                    print("WARNING: tagVisability must be an integer between 0 and 15, tag visability will be based on the number of chosen tags that can be tags in CD")
                else:
                    outputList.append("WARNING: tagVisability must be an integer between 0 and 15, tag visability will be based on the number of chosen tags that can be tags in CD")
    
        # 'ID' is used to match with the tags in the CD results file (can be 1-15)
        ID = 1
        
        # This variable is the output of this function
        # It stores the Tags and thresholds in 'tagDict' that can be Tags in CD
        tagDictOutput = {}
        
        # Loop through each tag that the user wants to use as a tag in CD 
        for tag, threshold in tagDict.items():
        
            # if the tag name is a column in the Excel file
            if tag in excelData.columns:
            
                # if the column in the Excel file is a boolean, integer, or float column
                if excelData.dtypes[tag] == "bool" or excelData.dtypes[tag] == "int64" or excelData.dtypes[tag] == "float64":
                
                    # change Tag Name and Description in CD database
                    cursor.execute("UPDATE DataDistributionBoxes SET Name = (?), Description = (?) WHERE BoxID = "+str(ID)+";", (tag, "Matching entry in: "+tag+".")) 
                    
                    # If the user has chosen a number of tags to be visable
                    if tagVisability is not None:
                        
                        # If the user wants the current tag to be visable
                        if ID <= tagVisability:
                            # Set the tags visability to True
                            cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'True' WHERE BoxID = "+str(ID)+";") 
                        
                        # If the user does not want the current tag to be visable
                        else:
                            # Set the tags visability to True
                            cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'False' WHERE BoxID = "+str(ID)+";") 
                    
                    # If the user has not chosen a number of tags to be visable
                    else:
                        # Set the tags visability to True
                        cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'True' WHERE BoxID = "+str(ID)+";") 
                    
                    # Add to ID, and add the current tag and threshold to the output dictionary
                    ID = ID + 1
                    tagDictOutput[tag] = threshold
                    
                # if the column in the Excel file is not a boolean, integer, or float column
                else:
                    if verbose:
                        print("WARNING: "+tag+" is not a boolean, integer, or float column")
                    else: 
                        outputList.append("WARNING: "+tag+" is not a boolean, integer, or float column")
                    
            # if the tag is not a column in the Excel file
            else:
                if verbose:
                    print("WARNING: "+tag+" not found in "+excelFilePath)
                else:
                    outputList.append("WARNING: "+tag+" not found in "+excelFilePath)
        
        # Loop through the tag IDs for tags that didn't have their names changed in CD
        while ID < 16:
        
            # If the user has chosen a number of tags to be visable
            if tagVisability is not None:
            
                # If the user wants the current tag to be visable
                if ID <= tagVisability:
                    # Set the tags visability to True
                    cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'True' WHERE BoxID = "+str(ID)+";") 
                
                # If the user does not want the current tag to be visable
                else:
                    # Set the tags visability to False
                    cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'False' WHERE BoxID = "+str(ID)+";") 
            
            # If the user has not chosen a number of tags to be visable
            else:
                # Set the tags visability to False
                cursor.execute("UPDATE DataDistributionBoxExtendedData SET ValueString = 'False' WHERE BoxID = "+str(ID)+";")
            ID = ID + 1
            
        # At least one of the chosen tags can be a tag in CD
        if tagDictOutput != {}:
            if verbose:
                return tagDictOutput
            else:
                return tagDictOutput, outputList
        
        # None of the chosen tags can be tags in CD
        else:
            if verbose:
                print("WARNING: none of the chosen Tags are boolean, integer, or float columns in "+excelFilePath)
                return None
            else:
                outputList.append("WARNING: none of the chosen Tags are boolean, integer, or float columns in "+excelFilePath)
                return None, outputList
    
    # Get info about errors
    except Exception as e:
        raise e
    
    
#####################################################################################
## Function: getTagBytes()
#####################################################################################
'''
This function gets the string stored in the Tags column of an Excel file and converts that 
string into bytes to be stored in the Compound Discoverer (CD) results file.

INPUT:
'tagString' = The string that needs to be converted to bytes.
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.

OUTPUT:
'tagbytes' = The tag bytes after being converted from a string.
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
            # 'keys' will be used to match rows in the SQLite statement to get the tag IDs
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
            
            # Create the tag byte string in the correct format for the CD results file
            # if an ID is found in the IDList, add \x01\x01 to indicate that the Tag with that ID is checked
            # \x00\x00 is used to idicate that the Tag with that ID is not checked
            # The tags ID can be 1-15
            for ID in range(1, 16):
                
                # If the current ID is not found in the ID list
                if binarySearch(ID, IDList) == -1:
                    tagBytes += b"\x00\x00"
                
                # If the current ID is found in the ID list
                else:
                    tagBytes += b"\x01\x01"
            
            return tagBytes
        
        # There is no tags
        else:
            return -1
    
    # Get info about errors
    except Exception as e:
        raise e
        
        
#####################################################################################
## Function: getTagString()
#####################################################################################
'''
This function gets the bytes stored in the Tags column of a Compound Discoverer (CD) results file and converts those 
bytes into a string to be stored in the Excel file if the Tag is set as visable. 

INPUT:
'tagBytes' = The bytes that needs to be converted to a string.
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.

OUTPUT:
'tagString' = The tag string after being converted from bytes.
'''

def getTagString(tagBytes, cdResultsFilePath, cursor):
    try:    
        # Get a list of IDs for the Tags that are visable in CD
        cursor.execute("SELECT BoxID FROM DataDistributionBoxExtendedData WHERE ValueString = 'True';")
        tempVisableIDList = cursor.fetchall()
        visableIDList = []
        for ID in tempVisableIDList:
            visableIDList.append(ID[0])
    
        # get the names of the Tags from CD
        cursor.execute("SELECT Name FROM DataDistributionBoxes WHERE BoxID < 16;")
        tempTagNameList = cursor.fetchall()
        tagNameList = []
        for name in tempTagNameList:
            tagNameList.append(name[0])
    
        # 'tagString' will be the output of this function
        tagString = ""
        # get the 'tagBytes' in the correct format for the following code
        tagBytes = str(tagBytes)
        tagBytes = tagBytes[2:-1]
    
        # Go through different sections of the 'tagBytes'
        # If a section == \x01\x01 then that Tag has been checked
        # and we can get the name for that tag from the 'tagNameList' and add that name to 'tagString'
        # The tags ID can be 1-15
        start = 0
        end = 8
        for ID in range(1, 16):
        
            # If the current tag is checked in CD
            if tagBytes[start : end] == "\\x01\\x01":
            
                # If the current tag is visable in CD
                if binarySearch(ID, visableIDList) != -1:
                    tagString = tagString + tagNameList[ID-1] +";"
            
            start = start + 8
            end = end + 8
                
        # Remove the last ';'
        if len(tagString) > 0:
            tagString = tagString[:-1]
        
        return tagString
    
    # Get info about errors
    except Exception as e:
        raise e


#####################################################################################
## Function: getColNameTupleListUpdatingInCD()
#####################################################################################
'''
This function gets a list of tuples.
Each tuple will contain a column DB name and a column Display name.
This function will get the correct columns for updating the Compound Discoverer (CD) Results file.

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.
'excelData' = A DataFrame containing Excel data.
'excelFilePath' = The path to an Excel file.
'oupdateColNameList' = A list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has already added to the CD results file).
'verbose' = Boolean value that controls the output to the console. 
    If False, hide outputs and return the outputs as a list.
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
    
OUTPUT:
'colNameTupleList' = A list of tuples. Each tuple contains a column's DB name and display name.
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
'''

# Gets the columns that are new, or editable
def getColNameTupleListUpdatingInCD(cdResultsFilePath, cursor, excelData, excelFilePath, updateColNameList, verbose, outputList):
    try:
        # Get the ID of the compound table in CD
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
                            
                        # If the column is not stored in the CD results file as the bytes type, or the column is the Tags column
                        cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                        if cursor.fetchall()[0][0] != 6 or colDisplayName == "Tags":
     
                            # If the column is editable
                            cursor.execute("SELECT Grid_AllowEdit FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                            if cursor.fetchall()[0][0] == 1:
                                        
                                # Get the column DBName and Display name, put those names in a tuple, then add that tuple to the list of tuples
                                cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                                colNameTuple = cursor.fetchall()[0]
                                colNameTupleList.append(colNameTuple)
                                        
                            # If the column is non-editable
                            else:
                                if verbose:
                                    print("WARNING: "+colDisplayName+" can't be updated because it is a non-editable column")
                                else:
                                    outputList.append("WARNING: "+colDisplayName+" can't be updated because it is a non-editable column")
                            
                        # If the column is stored in the CD results file as the bytes type
                        else:
                            if verbose:
                                print("WARNING "+colDisplayName+" can't be updated because of the way this column is stored in the database")
                            else:
                                outputList.append("WARNING "+colDisplayName+" can't be updated because of the way this column is stored in the database")
                            
                    # If the column is not in the CD results file
                    else:
                    
                        # convert the display name to a string that can be stored as an SQLite column name
                        colDBName = formatStringToSQLiteColumn(colDisplayName)
                                
                        # If the DB version of the column name is not already being used in the CD results file compound table
                        cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name=(?);", (colDBName, ))
                        if cursor.fetchall()[0][0] == 0:
                            
                            # Add DBName and Display name in a tuple, then add that tuple to the list of tuples
                            colNameTuple = (colDBName, colDisplayName)
                            colNameTupleList.append(colNameTuple)
                                    
                        # If the DB version of the column name is already being used in the CD results file
                        else:
                            if verbose:
                                print("WARNING: can't add "+colDisplayName+" to "+cdResultsFilePath+" because the SQLite friendly version of the name ("+colDBName+") is already being used")
                            else:
                                outputList.append("WARNING: can't add "+colDisplayName+" to "+cdResultsFilePath+" because the SQLite friendly version of the name ("+colDBName+") is already being used")
                       
                # If the column is not in the Excel file
                else:
                    if verbose:
                        print("WARNING: "+colDisplayName+" can't be found in "+excelFilePath)
                    else:
                        outputList.append("WARNING: "+colDisplayName+" can't be found in "+excelFilePath)
        
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
        
        if verbose:
            return colNameTupleList
        else:
            return colNameTupleList, outputList
    
    # Get info about errors
    except Exception as e:
        raise e
        

#####################################################################################
## Function: getColNameTupleListUpdatingInExcel()
#####################################################################################
'''
This function gets a list of tuples.
Each tuple will contain a column DB name and a column Display name.
This function will get the correct columns for updating the Excel file.

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.
'excelData' = A DataFrame containing Excel data.
'updateColNameList' = A list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file).
'verbose' = Boolean value that controls the output to the console. 
    If False, hide outputs and return the outputs as a list.
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
    
OUTPUT:
'colNameTupleList' = A list of tuples. Each tuple contains a column's DB name and display name.
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
'''

# Gets the columns that are new, or editable
def getColNameTupleListUpdatingInExcel(cdResultsFilePath, cursor, excelData, updateColNameList, verbose, outputList):
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
                        
                    # If the column is not stored in the CD results file as the bytes type, or is the Tags column
                    cursor.execute("SELECT CustomDataType FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                    if cursor.fetchall()[0][0] != 6 or colDisplayName == "Tags":
     
                        # Get the column DBName and Display name, put those names in a tuple, then add that tuple to the list of tuples
                        cursor.execute("SELECT DBColumnName, Property_DisplayName FROM DataTypesColumns WHERE DataTypeID = (?) AND Property_DisplayName = (?);", (compoundTblID, colDisplayName, ))
                        colNameTuple = cursor.fetchall()[0]
                        colNameTupleList.append(colNameTuple)
                            
                    # If the column is stored in the CD results file as the bytes type
                    else:
                        if verbose:
                            print("WARNING "+colDisplayName+" can't be updated because of the way this column is stored in the database")
                        else:
                            outputList.append("WARNING "+colDisplayName+" can't be updated because of the way this column is stored in the database")
                 
                # If the column is not in the CD results file
                else:
                    if verbose:
                        print("WARNING: "+colDisplayName+" can't be found in "+cdResultsFilePath)
                    else:
                        outputList.append("WARNING: "+colDisplayName+" can't be found in "+cdResultsFilePath)
    
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
        
        if verbose:
            return colNameTupleList
        else:
            return colNameTupleList, outputList
    
    # Get info about other errors
    except Exception as e:
        raise e
        
        
#####################################################################################
## Function: fillNAValuesInDF()
#####################################################################################
'''
This function fills the NA values in a dataframe. 
If a column is not the Tags column, then empty rows are filled with default values.

INPUT:
'excelData' = The DataFrame containing Excel data.

OUTPUT:
'excelData' = The DataFrame containing Excel data after NA values have been filled.
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
        raise e
        

#####################################################################################
## Function: addCompDiscIDsToExcel()
#####################################################################################
'''
This function adds the Compound Discoverer (CD) results file compound IDs to the Excel file.
Rows in the CD database and the Excel file are matched using the RetentionTime and MolecularWeight.
This is slow, but we only need to do this once. Matching rows will be faster once we have gotten the IDs.

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'cursor' = An SQLite cursor.
'excelRowCount' = The number of rows in the Excel Peak sheet.
'excelFilePath' = The path to an Excel file.
'excelData' = The dataframe containing Excel data.
'excelSheetName' = The name of the Excel sheet containing the peak data.
'verbose' = Boolean value that controls the output to the console. 
    If False, hide outputs and return the outputs as a list.
'outputList' = A list of outputs that have been hidden if 'verbose' is False.

OUTPUT:
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
'excelData' = The dataframe containing the Excel data, now with IDs.
'''

def addCompDiscIDsToExcel(cdResultsFilePath, cursor, excelRowCount, excelFilePath, excelData, excelSheetName, verbose, outputList):
    try:
        # list to hold the IDs that will get added to the Excel file
        IDList = []
                
        # Create a MolecularWeight & RetentionTime INDEX which improves performance with the upcoming SELECT statement
        cursor.execute("CREATE INDEX IF NOT EXISTS MW_RT ON ConsolidatedUnknownCompoundItems (MolecularWeight, RetentionTime);")
            
        # Loop through each row in the excel file
        for row in range(excelRowCount):
                    
            # Get molecular weight and retention time from the Excel data        
            # Molecular weight and retention time is needed to match rows between the excel file and CD results file
            MW = excelData.at[row,"Calc. MW"]
            RT = excelData.at[row,"RT [min]"]
                        
            # Get the ID of the current row and column from the CD results file
            # Round MolecularWeight to 5 decimal places and RetentionTime to 3 decimal places because those values are stored in that format in the Excel file
            cursor.execute("SELECT ID FROM ConsolidatedUnknownCompoundItems WHERE ROUND(MolecularWeight, 5) = ROUND((?), 5) and ROUND(RetentionTime, 3) = ROUND((?), 3);", (str(MW), str(RT), ))     
            selectStatementResults = cursor.fetchall()
            ID = selectStatementResults[0][0]
                    
            # If exactly one row in the CD results file matched with a row in the Excel file
            if len(selectStatementResults) == 1:
                # Add current ID to the ID list
                IDList.append(ID)
            
            # If multiple rows in the CD results file matched with a row in the Excel file
            elif len(selectStatementResults) > 1:
                if verbose: 
                    print("WARNING: multiple rows in "+cdResultsFilePath+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                else:
                    outputList.append("WARNING: multiple rows in "+cdResultsFilePath+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))

            # If no rows in the CD results file matched with a row in the Excel file
            else:
                if verbose: 
                    print("WARNING: no rows in "+cdResultsFilePath+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
                else:
                    outputList.append("WARNING: no rows in "+cdResultsFilePath+" have molecular weight = "+str(MW)+" and retention time = "+str(RT))
        
        # Drop the INDEX that was created earlier because it is not needed anymore   
        cursor.execute("DROP INDEX IF EXISTS MW_RT;")
                
        # Update the Excel dataframe with the ID list, save the dataframe to the Excel file, then return the Excel data with the IDs  
        IDData = pd.DataFrame (IDList, columns = ["CompDiscID"])
        excelData = pd.concat([IDData, excelData], axis=1)
        excelData.to_excel(excelFilePath, sheet_name=excelSheetName, index=False)
        
        if verbose:
            return excelData
        else:
            return excelData, outputList
        
    # If the Excel file doesn't have the correct sheet
    except ValueError:
        raise ValueError
        
    # If the Excel file can't be found
    except FileNotFoundError:
        raise FileNotFoundError
    
    # If permission to the Excel file was denied
    except PermissionError:
        raise PermissionError
        
    # Get info about errors
    except Exception as e:
        raise e
            

#####################################################################################
## Function: updateDataInCDResultsFile()
#####################################################################################
'''
This function imports data from an Excel file to a Compound Discoverer (CD) results file.

INPUT:
'cdResultsFilePath' = The path to a CD results file
'excelFilePath' = The path to an Excel file
'excelSheetName' = The name of the Excel sheet containing the peak data
'updateColNameList' = a list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file)
'optionalTagDict' = A dictionary. The keys of the dictionary are columns in the Excel file that the user wishes to use as Tags in CD,
    the values in the dictionary are thresholds that the user has chosen for checking the Tag boxes.
'tagVisability' = An integer from 0-15 setting the number Tags that the user wants to be visable in CD.
    Will contain a None value if the user hasn't chosen a number.

'verbose' = Boolean value that controls the output to the console. 
    If False, hide outputs and return the outputs as a list.
    
OUTPUT:    
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
'''
    
def updateDataInCDResultsFile(cdResultsFilePath, excelFilePath, excelSheetName, updateColNameList = None, optionalTagDict = None, tagVisability = None, verbose = True):
    
    if verbose: 
        print("\nImporting data from "+excelFilePath+" into "+cdResultsFilePath)
        outputList= None
    else:
        outputList = []
   
    try: 
        # If the results file can be found
        if os.path.exists(cdResultsFilePath):
            
            # Open connection to the Compound Discoverer File
            conn = sqlite3.connect(cdResultsFilePath)
            cursor = conn.cursor()
            
            # Get Excel data in a dataframe, fill NA values in that dataframe, and get the number of rows in that dataframe
            excelData = pd.read_excel(excelFilePath, sheet_name = excelSheetName)
            excelData = fillNAValuesInDF(excelData)
            excelRowCount = len(excelData.index)
    
            # If the user wants to update the Tags in CD using values from chosen Excel columns
            if optionalTagDict is not None:
                # Change the names of the Tags in CD based on the Excel column the user has chosen with 'optionalTagList'
                # Also get the user selected tags that are boolean, int, or float columns in the Excel file
                if verbose:
                    optionalTagDict = changeTagNamesAndVisability(cursor, excelData, excelFilePath, optionalTagDict, tagVisability, verbose, outputList)
                else:
                    optionalTagDict, outputList = changeTagNamesAndVisability(cursor, excelData, excelFilePath, optionalTagDict, tagVisability, verbose, outputList)
    
            # Get list of tuples
            # Each tuple will contains the column DB name and display name
            if verbose:
                colNameTupleList = getColNameTupleListUpdatingInCD(cdResultsFilePath, cursor, excelData, excelFilePath, updateColNameList, verbose, outputList)
            else:
                colNameTupleList, outputList = getColNameTupleListUpdatingInCD(cdResultsFilePath, cursor, excelData, excelFilePath, updateColNameList, verbose, outputList)

            # Get the ID of the compound table in the CD database
            cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
            compoundTblID = cursor.fetchall()[0][0]
        
            # Check if the compound table has the 'Cleaned' column
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
                # Check if the compound table has the 'OldCleaned' column
                cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name='OldCleaned';")
                
                # If the 'OldCleaned' column doesn't exist
                if cursor.fetchall()[0][0] == 0:
                
                    # Rename 'Cleaned' to 'OldCleaned'
                    cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems RENAME COLUMN Cleaned to OldCleaned")
                    
                    # Add 'OldCleaned' details to columns table
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
                
                    # Set 'Cleaned' to False for all rows
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET Cleaned = 'False';") 
                
                # The 'OldCleaned' column does exist    
                else:
                    # Copy current 'Cleaned' column into the 'OldCleaned' column
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET OldCleaned = Cleaned")
                    
                    # Set 'Cleaned' to False for all rows
                    cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET Cleaned = 'False';") 

            # If the Excel data doesn't contain the CD database IDs, add them
            if "CompDiscID" not in excelData.columns:
                if verbose:
                    excelData = addCompDiscIDsToExcel(cdResultsFilePath, cursor, excelRowCount, excelFilePath, excelData, excelSheetName, verbose, outputList) 
                else:
                    excelData, outputList = addCompDiscIDsToExcel(cdResultsFilePath, cursor, excelRowCount, excelFilePath, excelData, excelSheetName, verbose, outputList) 

            # Loop through each column tuple in the list of tuples, to update each column in the list
            for colNameTuple in colNameTupleList:

                # Get the column DB name and display name
                colDBName = colNameTuple[0]
                colDisplayName = colNameTuple[1]

                # If the column doesn't exist in the CD results file, we need to add the column before updating it
                cursor.execute("SELECT COUNT(*) AS CNTREC FROM pragma_table_info('ConsolidatedUnknownCompoundItems') WHERE name=(?);", (colDBName, ))
                if cursor.fetchall()[0][0] == 0:
                
                    # Get the data type of the Excel column
                    colDataType = excelData.dtypes[colDisplayName]
                           
                    # Add new column to CD results file and set default values based on the data type of the columns
                    # Also set customDataType and valueType which gets used in the CD database
                    if colDataType == "int64":
                        # Add column to compound table
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" INTEGER;")
                        # Add default values to the new column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", (str(0), ))
                                
                        # 'customeDataType' 2 is for int columns
                        customDataType = "2"
                        valueType = "A170C73A-BD79-493B-B24A-B981BAF6DCC5"
                                
                    elif colDataType == "float64":
                        # Add column to compound table
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" REAL;")
                        # Add default values to the new column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", (colDBName, str(0.0), ))
                        
                        # 'customeDataType' 3 is for float columns                        
                        customDataType = "3"
                        valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                    elif colDataType == "bool":
                        # Add column to compound table
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" NUMERIC;")
                        # Add default values to the new column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", ("False", ))
                        
                        # 'customeDataType' 4 is for string columns
                        # Storing bool values in CD as strings is easier
                        customDataType = "4"
                        valueType = "3245F562-3044-4BC0-9091-3813CA7AE5BC"
                                
                    else:
                        # Add column to compound table
                        cursor.execute("ALTER TABLE ConsolidatedUnknownCompoundItems ADD COLUMN "+colDBName+" TEXT;")
                        # Add default values to the new column
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?);", ("", ))
                        
                        # 'customeDataType' 4 is for string columns                        
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
                    
                    if verbose: 
                        print("Column: "+colDisplayName+" added to "+cdResultsFilePath)
                    else:
                        outputList.append("Column: "+colDisplayName+" added to "+cdResultsFilePath)

                # Loop through each row in the Excel file, to update the current column in the CD results file
                for row in range(excelRowCount):
                
                    # The Excel ID is needed to match rows between the excel file and CD results file
                    ID = excelData.at[row,"CompDiscID"]
                
                    # The Tags column needs to be handled differently
                    if colDBName == "Tags":
                    
                        # If the user wants to update the Tags in CD using values from chosen Excel columns
                        if optionalTagDict is not None:
                        
                            # Create a tagString containing Tags that pass the thresholds for the current row in the Excel file
                            tagString = "" 
                            for optionalTag, threshold in optionalTagDict.items():
                                tagValue = excelData.at[row,optionalTag]
                                if type(threshold) == bool:
                                    if tagValue == threshold:
                                        tagString = tagString + optionalTag + ";"
                                elif type(threshold) == int or type(threshold == float):
                                    if tagValue >= threshold:
                                        tagString = tagString + optionalTag + ";"
                                else:
                                    print("WARNING: Tags threshold should be a Boolean, integer, or float value")
                                    
                            # Remove the last ;
                            if len(tagString) > 0:
                                tagString = tagString[:-1]
                            
                            # Convert the tag string to bytes that CD can read
                            value = getTagBytes(tagString, cdResultsFilePath, cursor)
                        
                        # If the user wants to update the Tags in CD with a Tags string from Excel 
                        else:
                            # Convert the tag string to bytes that CD can read
                            value = getTagBytes(excelData.at[row,colDisplayName], cdResultsFilePath, cursor)
                            if value == -1:
                                value = b"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"
      
                        # Update the current row and column in the CD results file, also set Cleaned to True    
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?), Cleaned = 'True' WHERE ID = (?);", (value, str(ID), ))     
                                    
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
                        cursor.execute("UPDATE ConsolidatedUnknownCompoundItems SET "+colDBName+" = (?), Cleaned = 'True' WHERE ID = (?);", (str(value), str(ID), ))     

                if verbose: 
                    print("Column: "+colDisplayName+" updated")
                else:
                    outputList.append("Column: "+colDisplayName+" updated")
            
            # Save changes to CD database
            conn.commit()
            
            # Close the connection to the Compound Discoverer file
            cursor.close()
            conn.close()
            
            if verbose: 
                print(cdResultsFilePath+" updated")
            else:
                outputList.append(cdResultsFilePath+" updated")
                return outputList 
    
       # If the results file can't be found
        else:
            if verbose:
                print("ERROR: "+cdResultsFilePath+" can't be found")
            else:
                outputList.append("ERROR: "+cdResultsFilePath+" can't be found")
                return outputList

    # If the Excel file doesn't have the correct sheet
    except ValueError:
        if verbose:
            print("ERROR: Can't find "+excelSheetName+" in "+excelFilePath)
        else:
            outputList.append("ERROR: Can't find "+excelSheetName+" in "+excelFilePath)
            return outputList
        
    # If the Excel file can't be found
    except FileNotFoundError:
        if verbose:
            print("ERROR: Can't find "+excelFilePath)
        else:
            outputList.append("ERROR: Can't find "+excelFilePath)
            return outputList
    
    # If permission to the Excel file was denied
    except PermissionError:
        if verbose:
            print("ERROR: Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")
        else:
            outputList.append("ERROR: Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")
            return outputList
        
    # Get info about other errors
    except Exception as e:
        if verbose:
            logger.error(e, exc_info=True)
        else:
            outputList.append("ERROR: "+str(e))
            return outputList
            

#####################################################################################
## Function: updateDataInExcelFile()
#####################################################################################
'''
This function imports data from a Compound Discoverer (CD) results file into an Excel file.

INPUT:
'cdResultsFilePath' = The path to a CD results file.
'excelFilePath' = The path to an Excel file.
'excelSheetName' = The name of the Excel sheet containing the peak data.
'updateColNameList' = A list of columns in the Excel file that the user wishes to update (default is None), if this value is left as None, 
    all editable columns will be updated (Tags, Checked, Name, and any columns the user has added to the CD results file).
'optionalTagList' = The list of Boolean columns in the Excel file that the user has set as Tags/
'verbose' = Boolean value that controls the output to the console. 
    If False, hide outputs and return the outputs as a list.
    
OUTPUT:
'outputList' = A list of outputs that have been hidden if 'verbose' is False.
'''
    
def updateDataInExcelFile(cdResultsFilePath, excelFilePath, excelSheetName, updateColNameList = None, optionalTagList = None, verbose = True):
    
    if verbose == True: 
        print("\nImporting data from "+cdResultsFilePath+" into "+excelFilePath)
        outputList = None
    else:
        outputList = []
    
    try: 
        # If the results file can be found
        if os.path.exists(cdResultsFilePath):
        
            # Open connection to the Compound Discoverer File
            conn = sqlite3.connect(cdResultsFilePath)
            cursor = conn.cursor()
        
            # Get Excel data in a dataframe, fill NA values in that dataframe, and get the number of rows in that dataframe
            excelData = pd.read_excel(excelFilePath, sheet_name = excelSheetName)
            excelData = fillNAValuesInDF(excelData)
            excelRowCount = len(excelData.index)
            
            # Loop through each user chosen tag and use a temp list
            # to get the tags that are boolean columns in the Excel file
            # and store those tags back into 'optionalTagList'
            tempOptionalTagList = []
            for optionalTag in optionalTagList:
                
                # If the tag is a column in the Excel data
                if optionalTag in excelData.columns:
                    
                    # If the Excel column is a boolean column
                    if excelData.dtypes[optionalTag] == "bool":
                        tempOptionalTagList.append(optionalTag)
                    
                    # If the Excel column is not a boolean column
                    else:
                        if verbose == True:
                            print("WARNING: "+optionalTag+" isn't a boolean column")
                        else:
                            outputList.append("WARNING: "+optionalTag+" isn't a boolean column")
                
                # If the tag is not a column in the Excel data
                else:
                    if verbose == True:
                        print("WARNING: "+optionalTag+" can't be found in "+excelFilePath)
                    else:
                        outputList.append("WARNING: "+optionalTag+" can't be found in "+excelFilePath)
            
            # If at least one user chosen Tag is a boolean column in the Excel file 
            if tempOptionalTagList != []:
                optionalTagList = tempOptionalTagList
            
            # If none of the user chosen Tags are a boolean column
            else:
                optionalTagList = None
                if verbose == True:
                    print("WARNING: None of the chosen Tags were found in "+excelFilePath)
                else:
                    outputList.append("WARNING: None of the chosen Tags were found in "+excelFilePath)
            
            # Get list of tuples
            # Each tuple will contains the column DB name and display name
            if verbose:
                colNameTupleList = getColNameTupleListUpdatingInExcel(cdResultsFilePath, cursor, excelData, updateColNameList, verbose, outputList)
            else:
                colNameTupleList, outputList = getColNameTupleListUpdatingInExcel(cdResultsFilePath, cursor, excelData, updateColNameList, verbose, outputList)

            # Get the ID of compound table
            cursor.execute("SELECT DataTypeID FROM DataTypes WHERE TableName = 'ConsolidatedUnknownCompoundItems';")
            compoundTblID = cursor.fetchall()[0][0]
            
            # If the Excel data doesn't contain the CD database IDs, add them
            if "CompDiscID" not in excelData.columns:
                if verbose:
                    excelData = addCompDiscIDsToExcel(cdResultsFilePath, cursor, excelRowCount, excelFilePath, excelData, excelSheetName, verbose, outputList) 
                else:
                    excelData, outputList = addCompDiscIDsToExcel(cdResultsFilePath, cursor, excelRowCount, excelFilePath, excelData, excelSheetName, verbose, outputList)
 
            # Loop through each column tuple in the list of tuples, to update each column in the list
            for colNameTuple in colNameTupleList:
                
                # Get the column DB name and display name
                colDBName = colNameTuple[0]
                colDisplayName = colNameTuple[1]
            
                # If the column is not in the Excel file, add the column with placeholder values
                if colDisplayName not in excelData.columns:
                    excelData[colDisplayName] = [0]*excelRowCount
                
                    if verbose: 
                        print("Column: "+colDisplayName+" added to "+excelFilePath)
                    else:
                        outputList.append("Column: "+colDisplayName+" added to "+excelFilePath)
                
                # If the column is already in the Excel file
                else:
                    if verbose: 
                        print("Column: "+colDisplayName+" found in "+excelFilePath)
                    else:
                        outputList.append("Column: "+colDisplayName+" found in "+excelFilePath)
                     
                # Check if the column data type is boolean, the colIsBool variable gets used to make sure the column stays boolean
                colIsBool = False
                if excelData.dtypes[colDisplayName] == "bool":
                    colIsBool = True
                
                # Loop through each row in the excel file
                for row in range(excelRowCount):
                    
                    # The Excel ID is needed to match rows between the excel file and CD results file
                    ID = excelData.at[row,"CompDiscID"]
                        
                    # Get the value of the current row and column from the CD results file    
                    cursor.execute("SELECT "+colDBName+" FROM ConsolidatedUnknownCompoundItems WHERE ID = (?);", (str(ID), ))     
                    value = cursor.fetchall()[0][0]
                    
                    # The Tags column needs to be handled differently
                    if colDBName == "Tags":
                    
                        # Update the current row and column in the Excel data frame after converting the bytes value to a string
                        tagString = getTagString(value, cdResultsFilePath, cursor)
                        excelData.at[row,colDisplayName] = tagString
                        
                        # If the user has selected boolean tag columns in the Excel file to update
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
                        
                        # Update the current row and column in the Excel data
                        excelData.at[row,colDisplayName]=value

                # This is to make sure boolean values are set as bool in the Excel file
                if colIsBool:                 
                    excelData[colDisplayName] = excelData[colDisplayName].astype('bool')
            
            # Update the Excel file 
            excelData.to_excel(excelFilePath, sheet_name=excelSheetName, index=False)  
        
            # Close the connection to the Compound Discoverer file
            cursor.close()
            conn.close()
        
            if verbose: 
                print(excelFilePath+" updated")
            else:
                outputList.append(excelFilePath+" updated")
                return outputList
        
        # If the results file can't be found
        else:
            if verbose:
                print("ERROR: "+cdResultsFilePath+" can't be found")
            else:
                outputList.append("ERROR: "+cdResultsFilePath+" can't be found")
                return outputList

    # If the Excel file doesn't have the correct sheet
    except ValueError:
        if verbose:
            print("ERROR: Can't find "+excelSheetName+" in "+excelFilePath)
        else:
            outputList.append("ERROR: Can't find "+excelSheetName+" in "+excelFilePath)
            return outputList
        
    # If the Excel file can't be found
    except FileNotFoundError:
        if verbose:
            print("ERROR: Can't find "+excelFilePath)
        else:
            outputList.append("ERROR: Can't find "+excelFilePath)
            return outputList
    
    # If permission to the Excel file was denied
    except PermissionError:
        if verbose:
            print("ERROR: Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")
        else:
            outputList.append("ERROR: Couldn't gain permission to the Excel File. Make sure "+excelFilePath+" is not open in another program")
            return outputList
        
    # Get info about other errors
    except Exception as e:
        if verbose:
            logger.error(e, exc_info=True)
        else:
            outputList.append("ERROR: "+str(e))
            return outputList
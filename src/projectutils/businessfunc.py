import openpyxl
import os
import re

from constants.errorcodes import ERROR_SUCCESS,ERROR_FILE_NOT_FOUND
from projectutils.filefunc import openExcelFile, createTempFile

import msoffcrypto
import getpass

TEMPLATE_FILE_NAME = "test/TEMPLATE.xlsx"
TEMPLATE_FOLDER_NAME = "/Users/vipula/Documents/GitHub/PDF_OVERLAY/test/"
SOURCE_PATH = "test/"
TEMPLATE_SHEET_NAME = "Overlay"
RECORD_LIST_SHEET_NAME = "Data"
#Template column defs for overlays
TEMP_COL_INDEX = 0
TEMP_COL_NAME = 1
TEMP_COL_CONTENT = 2
TEMP_COL_LOC_X = 3
TEMP_COL_LOC_Y = 4
#Data content defs for overlays
TEMP_DATA_TYPE = 0
TEMP_DATA_FILE_NAME = 1
TEMP_DATA_IMD_TEXT = 1
#TEMP_DATA_FILE_LOCKED = 2
TEMP_DATA_FILE_SHEET = 2
TEMP_DATA_FILE_COL_KEY = 3
TEMP_DATA_FILE_COL_DATA = 4
#Template column defs for record IDs
REC_COL_INDEX = 0
REC_COL_KEY = 1 #Prinery Key
REC_COL_STR_ID = 2 #name



def loadTemplateData(templateFile,sheetName):
    """ Reads the text overlays from the template file.
    Args: templateFile (string): The template file name
          sheetName (string): The sheet name with data
    Returns: list: A list of ordered dictionaries """

    print("+ Fn: loadTemplateData")
    # variable to return data
    textOverlayList = [] 
    #check if the file path is valid
    if( not os.path.exists(templateFile)):
        print("Error [loadTemplateData]: Tempate file not found")
        return textOverlayList # error - empty
    #open template file, and load the sheet
    templateBook = openpyxl.load_workbook(templateFile)
    textOverlayDataSheet = templateBook[sheetName]
    print("Debug [loadTemplateData]: Sheet Size",textOverlayDataSheet.dimensions, textOverlayDataSheet.max_row, " Rows ", textOverlayDataSheet.max_column, " columns")
    #go through every text overlay item
    for overlays in textOverlayDataSheet.rows:
        #Get the index to a string. Empty cells will be string "None"
        rowIndex = str(overlays[TEMP_COL_INDEX].value)
        # stop if we reach an empty cell
        if("None" == rowIndex):
            break
        # the index has to be always numbers, skip others
        if( not rowIndex.isdigit()):
            print("Warning [loadTemplateData]: skip Row Index : ", rowIndex)
            continue
        dataString = str(overlays[TEMP_COL_CONTENT].value)
        #user initiated end of loop.
        if("None" == dataString):
            print("Warning [loadTemplateData]: User terminated at Index : ", rowIndex)
            break
        if(not (dataString.startswith('<') and dataString.endswith('>') and len(dataString) > 3)):
            print("Error [loadTemplateData]: Data Error at Index : ",rowIndex)
            break
        #process the file data. Get all the data points to a list
        data = re.findall(r'<(.*?)>',dataString)
        #check the item 2, File locked
        if "!T" == data[TEMP_DATA_TYPE]:
            print(rowIndex, "overlay type > immidiate text")
            #Save notmal text data
            textOverlayList.append({"name": overlays[TEMP_COL_NAME].value, 
                                    "text":{ "string": data[TEMP_DATA_IMD_TEXT], 
                                            "x": overlays[TEMP_COL_LOC_X].value, 
                                            "y": overlays[TEMP_COL_LOC_Y].value}})
        elif "!F" == data[0]:
            print(rowIndex, "overlay type > From file => ",data[TEMP_DATA_FILE_NAME])
            #isFileLocked = False
            #if data[TEMP_DATA_FILE_LOCKED] == "LOCKED=1":
            #    isFileLocked = True
            #save the extended data
            textOverlayList.append({"name": overlays[TEMP_COL_NAME].value, 
                                    "text":{ "string": None, 
                                            "x": overlays[TEMP_COL_LOC_X].value, 
                                            "y": overlays[TEMP_COL_LOC_Y].value},
                                    "file" : {"name": data[TEMP_DATA_FILE_NAME], 
                                                "sheet": data[TEMP_DATA_FILE_SHEET], 
                                                "primeryKey": data[TEMP_DATA_FILE_COL_KEY], 
                                                "value": data[TEMP_DATA_FILE_COL_DATA]}})
        else:
            print(rowIndex, "Error: undefined overlay type : ", data[0])
            break
    # Data store is done. return
    print("- Fn: loadTemplateData")
    return textOverlayList

def getFilesFromOverlayList(textOverlayList):
    """ Returns a list of file names in the overlay 
    Args: textOverlayList (list): list of directories
    Returns: list: A list of file names, strings"""

    print("+ Fn: getFilesFromOverlayList")
    fileNameList = []
    #go through each overlay
    for testOverlay in textOverlayList:
        if "file" in testOverlay:
            filedata = testOverlay["file"]
            filename = filedata["name"]
            print("Fould a file ", filename)
            if filename not in fileNameList:
                fileNameList.append(filename)
    print("- Fn: getFilesFromOverlayList")
    return fileNameList


def loadRecordIdList(templateFile,sheetName):
    """ loadRecordIdList: This will load the records to process """
    print("+ Fn: loadRecordIdList")
    # variable to return data
    recordIdList = [] 
    #check if the file path is valid
    if( not os.path.exists(templateFile)):
        print("Error [loadTemplateData]: Tempate file not found")
        return recordIdList # error - empty
    #open template file, and load the sheet
    templateBook = openpyxl.load_workbook(templateFile)
    recordIdDataSheet = templateBook[sheetName]
    print("Debug [loadTemplateData]: Sheet Size",recordIdDataSheet.dimensions, recordIdDataSheet.max_row, " Rows ", recordIdDataSheet.max_column, " columns")
    #go through every text overlay item
    for record in recordIdDataSheet.rows:
        #Get the index to a string. Empty cells will be string "None"
        rowIndex = str(record[REC_COL_INDEX].value)
        print(rowIndex)
        # stop if we reach an empty cell
        if("None" == rowIndex):
            break
        # the index has to be always numbers, skip others
        if( not rowIndex.isdigit()):
            print("Warning [loadTemplateData]: skip Row Index : ", rowIndex)
            continue
        primeryKey = str(record[REC_COL_KEY].value)
        #user initiated end of loop.
        if("None" == primeryKey):
            print("Warning [loadTemplateData]: User terminated at Index : ", rowIndex)
            break
        if not primeryKey.isdigit():
            print("Error [primeryKey]: Data Error at Index : ",rowIndex)
            break
        recordIdList.append({"key": int(primeryKey), "identifier": str(record[REC_COL_STR_ID].value)})
    return recordIdList


def openSourceFiles(fileNameList):
    """openSourceFiles"""
    print("+Fn openSourceFiles, PARAM = ", fileNameList)
    fileObjectList = []
    for fileName in fileNameList:
        fileObjectList.append({"name": fileName, "object": None, "tempFile": None})
    return fileObjectList


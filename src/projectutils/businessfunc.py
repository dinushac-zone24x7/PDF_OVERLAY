import openpyxl
import os
import re

from constants.errorcodes import ERROR_SUCCESS, ERROR_FILE_NOT_FOUND, ERROR_TEMP_FILE_DELETE
from projectutils.filefunc import openExcelFile, createTempFile, removeFiles

import msoffcrypto
import getpass

TEMPLATE_FILE_NAME = "test/TEMPLATE.xlsx"
TEMPLATE_FOLDER_NAME = "/Users/vipula/Documents/GitHub/PDF_OVERLAY/test/"
SOURCE_PATH = "test/"
TEMPLATE_SHEET_NAME = "Overlay"
RECORD_LIST_SHEET_NAME = "Data"
#Template column defs
TEMP_COL_INDEX = 0
TEMP_COL_NAME = 1
TEMP_COL_CONTENT = 2
TEMP_COL_LOC_X = 3
TEMP_COL_LOC_Y = 4
#Data content defs
TEMP_DATA_TYPE = 0
TEMP_DATA_FILE_NAME = 1
TEMP_DATA_IMD_TEXT = 1
TEMP_DATA_FILE_LOCKED = 2
TEMP_DATA_FILE_SHEET = 3
TEMP_DATA_FILE_COL_KEY = 4
TEMP_DATA_FILE_COL_DATA = 5

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
            isFileLocked = False
            if data[TEMP_DATA_FILE_LOCKED] == "LOCKED=1":
                isFileLocked = True
                #save the extended data
                textOverlayList.append({"name": overlays[TEMP_COL_NAME].value, 
                                        "text":{ "string": None, 
                                                "x": overlays[TEMP_COL_LOC_X].value, 
                                                "y": overlays[TEMP_COL_LOC_Y].value},
                                        "file" : {"name": data[TEMP_DATA_FILE_NAME], 
                                                  "isLocked": isFileLocked, 
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


def loadRecordIdList(templateFileName):
    """DUMMY FUNC <TO DO>"""
    print("+Fn loadRecordIdList = DUMMY, PARAM = ", templateFileName)
    recordIdList = [10,23,35,135]
    return recordIdList


def openSourceFiles(fileList):
    """DUMMY FUNC <TO DO>"""
    print("+Fn openSourceFiles = DUMMY, PARAM = ", fileList)
    fileObjectList = []
    fileObjectList.append({"name": "name", "object": None})
    return fileObjectList

def unused():
    #get the files to be used as source 
    SourceFileList = getFilesFromOverlayList(None)
    print(SourceFileList)

    #open the files
    fileModule = []
    tempFileList = []
    fileObject = None

    for sourceFile in SourceFileList:
        tempFileIndex = 1
        sourceFileFullPath = SOURCE_PATH+sourceFile
        try:
            with open (sourceFileFullPath, 'rb') as currentFile:
                office_file = msoffcrypto.OfficeFile(currentFile)
                if office_file.is_encrypted():
                    print("Fn Main: File Encrypted.")
                    #get the password
                    password = getpass.getpass("Enter password for "+ sourceFile + ": ")
                    tempFileFullPath = SOURCE_PATH+"~temp~"+str(tempFileIndex)+".xlsx"
                    tempFileIndex = tempFileIndex + 1
                    #create a temp file
                    if( ERROR_SUCCESS != createTempFile(sourceFileFullPath,password,tempFileFullPath)):
                        print("Fn: Main Error exit")
                        exit(ERROR_FILE_NOT_FOUND)
                    #add the file name to temp file list.
                    tempFileList.append({"name":tempFileFullPath, "delete": True})
                    fileObject = openExcelFile(tempFileFullPath)
                else:
                    fileObject = openExcelFile(sourceFileFullPath)
        except Exception as e:
            print(f"An error occurred: {e}")
        #add the file object to the list
        fileModule.append({"name": sourceFile, "fileObject": fileObject})


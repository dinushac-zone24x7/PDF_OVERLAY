""" File Functions 
    ssrc/projectutils/filefunc.py 
    This file contains the functions related file manipulations
    Author: vipulasrilanka@yahoo.com 
    (c) 2024 """

import msoffcrypto
import openpyxl
import os
from constants.errorcodes import ERROR_FILE_NOT_FOUND, ERROR_SUCCESS, ERROR_FILE_ENCRYPTED, ERROR_UNKNOWN

# passwords perdata and saldata 


def createTempFile(sourceFileName,password,tempFileName):
    """ create a temp data file from a locked excel file"""
    print("Fn: createTempFile", sourceFileName,password,tempFileName)
    #check if the source file is there
    if(not os.path.exists(sourceFileName)):
        print("Error [createTempFile]: Source file not found")
        return ERROR_FILE_NOT_FOUND # error 
    with open(sourceFileName, 'rb') as file:
        officeFile = msoffcrypto.OfficeFile(file)
        officeFile.load_key(password=password)  # Provide the password
        # Decrypt the file and save it as a temporary file
        with open(tempFileName, 'wb') as decryptedFile:
            officeFile.decrypt(decryptedFile)
    return ERROR_SUCCESS

def openExcelFile(sourceFileName):
    """ Open an excel file and return the file object"""
    print ("+Fn: openSourceFile",sourceFileName)
    if(not os.path.exists(sourceFileName)):
        print("Error [openExcelFile]: Source file not found")
        return {"error": ERROR_FILE_NOT_FOUND, "object": None} # error
    else:
        try:
            workbook = openpyxl.load_workbook(sourceFileName, data_only=True)
        except Exception as e:
            print(f"Error: {e}")
            #check if this is a password protected file
            with open (sourceFileName, 'rb') as excelFile:
                officeFile = msoffcrypto.OfficeFile(excelFile)
                if officeFile.is_encrypted():
                    print("Fn [openExcelFile]:: File Encrypted.")
                    return {"error": ERROR_FILE_ENCRYPTED, "object": None}
                else:
                    #This is an unknown error
                    print("Error [openExcelFile]: unknown Error")
                    return {"error": ERROR_UNKNOWN, "object": None} # error
    return {"error": ERROR_SUCCESS, "object": workbook}

def removeFiles(fileList):
    """Remove Files
    Args: fileList (list) : Dictionary of temp file name list
    Returns: int : error count"""
    print("+Fn: removeFiles")
    errorCount = 0
    for files in fileList:
        if(files["delete"] & os.path.exists(files["name"])):
            os.remove(files["name"])
            print(f"Temporary files deleted.")
        else:
            print(f"ERROR: File not found.")
            print(files["name"])
            errorCount = errorCount + 1
    return errorCount

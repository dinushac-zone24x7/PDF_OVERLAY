import msoffcrypto
import openpyxl
import os
from constants.errorcodes import ERROR_FILE_NOT_FOUND, ERROR_SUCCESS

# passwords perdata and saldata 


def createTempFile(sourceFileName,password,tempFileName):
    """ create a temp data file from a locked excel file"""
    print("Fn: createTempFile", sourceFileName,password,tempFileName)
    #check if the source file is there
    if( not os.path.exists(sourceFileName)):
        print("Error [createTempFile]: Source file not found")
        return ERROR_FILE_NOT_FOUND # error 
    with open(sourceFileName, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)  # Provide the password
        # Decrypt the file and save it as a temporary file
        with open(tempFileName, 'wb') as decrypted_file:
            office_file.decrypt(decrypted_file)
    return ERROR_SUCCESS

def openExcelFile(sourceFileName):
    """ Open an excel file and return the file object"""
    print ("+Fn: openSourceFile",sourceFileName)
    workbook = openpyxl.load_workbook(sourceFileName)
    return workbook

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

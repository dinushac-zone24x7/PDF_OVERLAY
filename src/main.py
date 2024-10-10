import msoffcrypto
import getpass
from projectutils.businessfunc import loadTemplateData, getFilesFromOverlayList
from projectutils.filefunc import openExcelFile, createTempFile, removeFiles
from constants.errorcodes import ERROR_SUCCESS, ERROR_FILE_NOT_FOUND, ERROR_TEMP_FILE_DELETE

TEMPLATE_FILE_NAME = "test/TEMPLATE.xlsx"
SOURCE_PATH = "test/"

# main program

#load the template
textOverlayList =  loadTemplateData(TEMPLATE_FILE_NAME,"Overlay")

#get the files to be used as source 
SourceFileList = getFilesFromOverlayList(textOverlayList)
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

print(tempFileList)
if(0 != removeFiles(tempFileList)):
    exit(ERROR_TEMP_FILE_DELETE)
else:
    exit(ERROR_SUCCESS)
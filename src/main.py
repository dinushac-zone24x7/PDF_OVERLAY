import threading
import time
from projectutils.guifunc import showStatus, getExcelFileName, getPassword  # Import GUI functions
from projectutils.guifunc import WINDOW_QUIT # Import GUI constants, Window
from projectutils.guifunc import MESSAGE_NEW, MESSAGE_ADD, MESSAGE_CLEAR # Import GUI constants Message
from projectutils.guifunc import GET_PASSWORD,RETURN_PASSWORD # Import GUI constants password
from projectutils.businessfunc import loadTemplateData, getFilesFromOverlayList, openSourceFiles, loadRecordIdList
from projectutils.businessfunc import TEMPLATE_SHEET_NAME, TEMPLATE_FOLDER_NAME, RECORD_LIST_SHEET_NAME
from projectutils.filefunc import openExcelFile, createTempFile, removeFiles
from constants.errorcodes import ERROR_FILE_NOT_FOUND, ERROR_SUCCESS, ERROR_FILE_ENCRYPTED, ERROR_UNKNOWN



def update_message(message_holder, action, message, isResetId):
    """Update the message in the holder (first element of the list)."""
    if isinstance(message, str):
        print(message)
    if isResetId:
        message_holder["id"] = 0
    else:
        message_holder["id"] = message_holder["id"] + 1
    message_holder["action"] = action
    message_holder["message"] = message

def processRecord(message_holder,FileObjectList,recordID,textOverlayList):
    """Simulate background message changes."""
    print("+Fn processRecord : ",recordID)
    update_message(message_holder, MESSAGE_NEW, "Status updated: Start...",False)
    for textOverlay in textOverlayList:
        print (textOverlay)
    update_message(message_holder, WINDOW_QUIT, None,False)  # This should close the status window

def noneed(message_holder):
    time.sleep(2)
    update_message(message_holder, GET_PASSWORD, "Password for File X",False)
    while(1):
        time.sleep(0.5)
        print(".", message_holder["action"])
        if RETURN_PASSWORD == message_holder["action"]:
            break
    print ("password = ", message_holder["message"])
    update_message(message_holder, MESSAGE_ADD, "Status updated: Processing...",False)
    time.sleep(1)
    update_message(message_holder, MESSAGE_ADD, "Status updated: Almost done...",False)
    time.sleep(1)
    update_message(message_holder, WINDOW_QUIT, None,False)  # This should close the status window

def main():
    """Main Function"""
    #get the template file name
    templateFileName = getExcelFileName("Open Template",TEMPLATE_FOLDER_NAME)
    #get Recoed ID list
    recordIDList = loadRecordIdList(templateFileName,RECORD_LIST_SHEET_NAME)
    #get the overlay list
    textOverlayList = loadTemplateData(templateFileName,TEMPLATE_SHEET_NAME)
    #get the file list
    fileNameList = getFilesFromOverlayList(textOverlayList)
    #get File Object list
    tempFileList = []
    fileObjectList = []
    for sourceFile in fileNameList:
        tempFileIndex = 1
        sourceFileFullPath = sourceFile
        while(1):
            returnValue = openExcelFile(sourceFileFullPath)
            if ERROR_SUCCESS == returnValue["error"]:
                fileObjectList.append({"name": sourceFile, "object": returnValue["object"]})
                break
            elif ERROR_FILE_NOT_FOUND == returnValue["error"]:
                sourceFileFullPath = getExcelFileName(f"Open {sourceFile}","./")
                continue
            elif ERROR_FILE_ENCRYPTED == returnValue["error"]:
                password = getPassword(sourceFile)
                tempFileName = "~temp~"+str(tempFileIndex)+".xlsx"
                if ERROR_SUCCESS != createTempFile(sourceFileFullPath,password,tempFileName):
                    print("Fn: Main => Can not create file. unknown Error")
                    return ERROR_UNKNOWN
                sourceFileFullPath = tempFileName
                tempFileList.append(tempFileName)
                tempFileIndex = tempFileIndex + 1
                continue
            else:
                return ERROR_UNKNOWN

    for recordId in recordIDList:
        message_holder = {"id": 0, "action": MESSAGE_CLEAR, "message": None}
        #update_message(message_holder,MESSAGE_NEW,templateFileName,True)
        windowName = "Status of Record ID = [" + str(recordId) + "]"
        # Start a background thread to simulate message updates
        thread = threading.Thread(target=processRecord, args=(message_holder,fileObjectList,recordId,textOverlayList,))
        thread.daemon = True  # Daemon thread will close with the main program
        thread.start()
        # Run the Tkinter GUI (must run in the main thread)
        showStatus(message_holder, windowName)
    
    removeFiles(tempFileList)

if __name__ == "__main__":
    main()

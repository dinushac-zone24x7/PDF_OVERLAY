import threading
import time
from projectutils.guifunc import showStatus, getFileName, getPassword  # Import GUI functions
from projectutils.guifunc import MESSAGE_QUIT, MESSAGE_NEW, MESSAGE_ADD, MESSAGE_CLEAR   # Import GUI constants
from projectutils.businessfunc import loadTemplateData, getFilesFromOverlayList, openSourceFiles, loadRecordIdList
from projectutils.businessfunc import TEMPLATE_SHEET_NAME, TEMPLATE_FOLDER_NAME, RECORD_LIST_SHEET_NAME

def update_message(message_holder, action, message, isResetId):
    if isinstance(message, str):
        print(message)
    """Update the message in the holder (first element of the list)."""
    if isResetId:
        message_holder["id"] = 0
    else:
        message_holder["id"] = message_holder["id"] + 1
    message_holder["action"] = action
    message_holder["message"] = message

def processRecord(message_holder,FileNameList,recordID):
    print("+Fn processRecord : ",recordID, FileNameList)
    """Simulate background message changes."""
    update_message(message_holder, MESSAGE_NEW, "Status updated: Start...",False)
    time.sleep(2)
    update_message(message_holder, MESSAGE_ADD, "Status updated: Processing...",False)
    time.sleep(1)
    update_message(message_holder, MESSAGE_ADD, "Status updated: Almost done...",False)
    time.sleep(1)
    update_message(message_holder, MESSAGE_QUIT, None,False)  # This should close the status window

def main():
    #get the template file name
    templateFileName = None
    templateFileName = getFileName(TEMPLATE_FOLDER_NAME)
    #get Recoed ID list
    recordIDList = loadRecordIdList(templateFileName)
    #get the overlay list
    #textOverlayList = loadTemplateData(templateFileName,TEMPLATE_SHEET_NAME)
    #get the file list
    #fileNameList = getFilesFromOverlayList(textOverlayList)
    #get File Object list
    #fileObjectList = openSourceFiles(fileNameList)
    fileObjectList = []

    for recordId in recordIDList:
        message_holder = {"id": 0, "action": MESSAGE_CLEAR, "message": None}
        #update_message(message_holder,MESSAGE_NEW,templateFileName,True)
        windowName = "Status of Record ID = [" + str(recordId) + "]"
        # Start a background thread to simulate message updates
        thread = threading.Thread(target=processRecord, args=(message_holder,fileObjectList,recordId,))
        thread.daemon = True  # Daemon thread will close with the main program
        thread.start()
        # Run the Tkinter GUI (must run in the main thread)
        showStatus(message_holder, windowName)

if __name__ == "__main__":
    main()

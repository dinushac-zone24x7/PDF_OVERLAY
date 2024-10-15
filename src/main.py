import threading
import time
from projectutils.guifunc import showStatus, getExcelFileName, getPassword, getPdfFileName  # Import GUI functions
from projectutils.guifunc import WINDOW_QUIT # Import GUI constants, Window
from projectutils.guifunc import MESSAGE_NEW, MESSAGE_ADD, MESSAGE_CLEAR # Import GUI constants Message
from projectutils.guifunc import GET_PASSWORD,RETURN_PASSWORD # Import GUI constants password
from projectutils.businessfunc import loadTemplateData, getFilesFromOverlayList, openSourceFiles, loadRecordIdList
from projectutils.businessfunc import TEMPLATE_SHEET_NAME, TEMPLATE_FOLDER_NAME, RECORD_LIST_SHEET_NAME
from projectutils.filefunc import openExcelFile, createTempFile, removeFiles
from constants.errorcodes import ERROR_FILE_NOT_FOUND, ERROR_SUCCESS, ERROR_FILE_ENCRYPTED, ERROR_UNKNOWN
from projectutils.pdfFunc import addOverlayToPdf

def update_message(messageHolder, action, message, isResetId):
    """Update the message in the holder (first element of the list)."""
    if isinstance(message, str):
        print(message)
    if isResetId:
        messageHolder["id"] = 0
    else:
        messageHolder["id"] = messageHolder["id"] + 1
    messageHolder["action"] = action
    messageHolder["message"] = message

def processRecord(messageHolder,FileObjectList,recordID,textOverlayList,PdfTemplateName, outputFileName):
    """Simulate background message changes."""
    print("+Fn processRecord : ",recordID)
    pdfOverlayList= []
    update_message(messageHolder, MESSAGE_NEW, "Status updated: Start...",False)
    #time.sleep(1)
    for textOverlay in textOverlayList:
        overlayName = textOverlay["name"]
        print (overlayName)
        if overlayName.startswith("!<CONCAT>"):
            #concatnate
            print("* => Concatnate")
        else:
            overlayText = textOverlay["text"]
            ovelayLocX = overlayText["x"]
            ovelayLocY = overlayText["y"]
            overlayString = overlayText["string"]
            if not None == overlayString:
                print("* => Immidiate string")
                pdfOverlayList.append({"text":overlayText, "x": ovelayLocX, "y": ovelayLocY})
            else:
                print ("* => From File")
    update_message(messageHolder, MESSAGE_ADD, "Creating PDF File ",False)
    addOverlayToPdf(PdfTemplateName, outputFileName, pdfOverlayList)
    print("created PDF", outputFileName)
    #time.sleep(1)
    update_message(messageHolder, MESSAGE_ADD, "Done..! ",False)
    #time.sleep(1)
    update_message(messageHolder, WINDOW_QUIT, None,False)  # This should close the status window
    #time.sleep(1)

def main():
    """Main Function"""
    pdfFileName = getPdfFileName("Select PDF file",TEMPLATE_FOLDER_NAME)
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
    tempFileIndex = 1 #Keep track on files
    for sourceFile in fileNameList:
        sourceFileFullPath = sourceFile
        while(1):
            returnValue = openExcelFile(sourceFileFullPath)
            if ERROR_SUCCESS == returnValue["error"]:
                fileObjectList.append({"name": sourceFile, "object": returnValue["object"]})
                break
            elif ERROR_FILE_NOT_FOUND == returnValue["error"]:
                sourceFileFullPath = getExcelFileName(f"Open {sourceFile}",TEMPLATE_FOLDER_NAME)
                continue
            elif ERROR_FILE_ENCRYPTED == returnValue["error"]:
                password = getPassword(sourceFile)
                tempFileName = "~temp~"+str(tempFileIndex)+".xlsx"
                if ERROR_SUCCESS != createTempFile(sourceFileFullPath,password,tempFileName):
                    print("Fn: Main => Can not create file. unknown Error")
                    return ERROR_UNKNOWN
                sourceFileFullPath = tempFileName
                tempFileList.append({"delete":True, "name": tempFileName})
                tempFileIndex = tempFileIndex + 1
                continue
            else:
                return ERROR_UNKNOWN

    for recordId in recordIDList:
        messageHolder = {"id": 0, "action": MESSAGE_CLEAR, "message": None}
        #update_message(messageHolder,MESSAGE_NEW,templateFileName,True)
        windowName = "Status of Record ID = [" + str(recordId["identifier"]) + "]"
        outPutFileName = str(recordId["identifier"])+"-"+str(recordId["key"])+".pdf"
        # Start a background thread to simulate message updates
        thread = threading.Thread(target=processRecord, args=(messageHolder,fileObjectList,recordId,textOverlayList,pdfFileName,outPutFileName))
        thread.daemon = True  # Daemon thread will close with the main program
        thread.start()
        # Run the Tkinter GUI (must run in the main thread)
        #showStatus(messageHolder, windowName) #disable for now. There is a bug/unknown
        while thread.is_alive():
            print("Thread is still running...")
            time.sleep(1)
        print("Thread has finished.")
    #we are done, so delete the temp files.  
    removeFiles(tempFileList)
    return ERROR_SUCCESS

def TestAlgo():
    pdfOverlayList = []
    templateFileName = getExcelFileName("Open Template",TEMPLATE_FOLDER_NAME)
    #get the overlay list
    textOverlayList = loadTemplateData(templateFileName,TEMPLATE_SHEET_NAME)
    for textOverlay in textOverlayList:
        overlayName = textOverlay["name"]
        print (overlayName)
        if overlayName.startswith("!<CONCAT>"):
            #concatnate
            print("* => Concatnate")
        else:
            overlayText = textOverlay["text"]
            ovelayLocX = overlayText["x"]
            ovelayLocY = overlayText["y"]
            overlayString = overlayText["string"]
            if not None == overlayString:
                print("* => Immidiate string")
                pdfOverlayList.append({"text":overlayText, "x": ovelayLocX, "y": ovelayLocY})
            else:
                print ("* => From File")
    # We are done.. exit now.
    return ERROR_SUCCESS


if __name__ == "__main__":
    main()
#    TestAlgo()

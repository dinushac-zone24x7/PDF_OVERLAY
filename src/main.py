import threading
import time
import sys
import os
from constants.templatedata import TEMPLATE_SHEET_NAME, TEMPLATE_FOLDER_NAME, RECORD_LIST_SHEET_NAME
from constants.errorcodes import ERROR_FILE_NOT_FOUND, ERROR_SUCCESS, ERROR_FILE_ENCRYPTED, ERROR_UNKNOWN, ERROR_ITEM_NOT_FOUND, ERROR_GENERAL_FAILIURE
from constants.pdfData import PDF_FIRST_PAGE, getPdfPage
from projectutils.guifunc import WINDOW_QUIT # Import GUI constants, Window
from projectutils.guifunc import MESSAGE_NEW, MESSAGE_ADD, MESSAGE_CLEAR # Import GUI constants Message
from projectutils.guifunc import showStatus, getExcelFileName, getPassword, getPdfFileName  # Import GUI functions
from projectutils.businessfunc import loadTemplateData, getFilesFromOverlayList, loadRecordIdList
from projectutils.businessfunc import getStringFromFileObject, concatString
from projectutils.filefunc import openExcelFile, createTempFile, removeFiles
from projectutils.filefunc import saveSessionData, loadSessionData
from projectutils.pdfFunc import addOverlayToPdf

def getSessionData(argv):
    """Get the session data including source file names"""
    sessionData = {"error": ERROR_SUCCESS, "rootFolder":None, "sessionFileName": None, "pdfFileName": None,"templateFileName": None,"sourceFiles": []}
    sessionData["rootFolder"] = os.path.dirname(os.path.abspath(argv[0]))
    if len(argv) < 2:
        #no args, Get sessoin data manually
        #get the PDF template file name from user
        sessionData["pdfFileName"] = getPdfFileName("Select PDF file",sessionData["rootFolder"])
        #get the template file name from user
        sessionData["templateFileName"] = getExcelFileName("Open Template",sessionData["rootFolder"])
    else:
        sessionData["sessionFileName"] = "argv[1]"
        print("Using Session File => ", sessionData["sessionFileName"])
        savedSession = loadSessionData(sessionData["sessionFileName"])
        if isinstance(savedSession,dict):
            sessionData["pdfFileName"] = savedSession["pdfFileName"] 
            sessionData["templateFileName"] = savedSession["templateFileName"]
            sessionData["sourceFiles"] = savedSession["sourceFiles"]
        else:
            sessionData["error"] = ERROR_UNKNOWN
    return sessionData

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


def processRecord(messageHolder,FileObjectList,recordID,textOverlayList,PdfTemplateName, PdfTemplatePage, outputFileName):
    """ Main record proccesor thread.
        This will each overlay for a given record 
    Args:   messageHolder (directory): contains the process status
            FileObjectList(list): list of directories with file names and mapping objects
            recordID (int): Record ID
            textOverlayList (list): directories with ordered overlay list
            PdfTemplateName (string) name of the PDF template
            outputFileName (string) name of the output file"""
    print("+Fn processRecord : ",recordID["identifier"])
    pdfOverlayList= []
    update_message(messageHolder, MESSAGE_NEW, "Status updated: Start...",False)
    for textOverlay in textOverlayList:
        overlayName = textOverlay["name"]
        if None == overlayName:
            print("Error - Bad overlay")
            exit (ERROR_UNKNOWN)
        overlayText = textOverlay["text"]
        ovelayLocX = overlayText["x"]
        ovelayLocY = overlayText["y"]
        overlayString = overlayText["string"]
        if None == overlayString:
            #Not an immidiate string
            print("Not Immidiate string")
            fileInfo = textOverlay["file"]
            fileName = fileInfo["name"]
            fileSheetName = fileInfo["sheet"]
            primeryKeyCol = fileInfo["primeryKey"]
            valueCol = fileInfo["value"]
            overlayString = getStringFromFileObject(fileName,FileObjectList,fileSheetName,recordID["key"],primeryKeyCol,valueCol)
            print("Type of overlayString = str ? ",isinstance(overlayString, str))
            if not isinstance(overlayString, str):
                # Error returned by the function
                print("ERROR: Can not find the primery key.!")
                return (ERROR_ITEM_NOT_FOUND)
        else:
            #an immidiate string
            print("Immidiate string")
        if overlayName.startswith("!<CONCAT>"):
            #concatnate
            print("* => Concatnate")
            concatString(pdfOverlayList,overlayName,overlayString)
        else:
            print("overlayString ["+ overlayName +"] = "+ str(overlayString), ovelayLocX, ovelayLocY)
            pdfOverlayList.append({"name":overlayName,"string":overlayString, "x": ovelayLocX, "y": ovelayLocY})

    update_message(messageHolder, MESSAGE_ADD, "Creating PDF File ",False)
    addOverlayToPdf(PdfTemplateName, PdfTemplatePage, outputFileName, pdfOverlayList)
    print("created PDF", outputFileName)
    update_message(messageHolder, MESSAGE_ADD, "Done..! ",False)
    update_message(messageHolder, WINDOW_QUIT, None,False)  # This should close the status window
    return ERROR_SUCCESS

def main():
    """Main Function"""
    #create place holders for session variables
    sessionData = getSessionData(sys.argv)
    if not sessionData["error"] == ERROR_SUCCESS:
        print("ERROR: Can not load session data.")
        exit(ERROR_GENERAL_FAILIURE)
    #get Recoed ID list
    recordIDList = loadRecordIdList(sessionData["templateFileName"],RECORD_LIST_SHEET_NAME)
    #get the overlay list
    textOverlayList = loadTemplateData(sessionData["templateFileName"],TEMPLATE_SHEET_NAME)
    #check errors and exut
    if isinstance(textOverlayList,int) or  isinstance(recordIDList,int):
        print("ERROR: Check the template file")
        exit(ERROR_GENERAL_FAILIURE)
    if isinstance(recordIDList, list) and len(recordIDList) < 1:
        # there are no records to process
        print("ERROR: can not find records to process. Check the template file")
        exit(ERROR_GENERAL_FAILIURE)
    #get the file list
    fileNameList = getFilesFromOverlayList(textOverlayList)
    #get File Object list
    tempFileList = []
    fileObjectList = []
    tempFileIndex = 1 #Keep track on files
    for sourceFile in fileNameList:
        sourceFileFullPath = os.path.join(sessionData["rootFolder"],sourceFile)
        while(1):
            returnValue = openExcelFile(sourceFileFullPath)
            if ERROR_SUCCESS == returnValue["error"]:
                fileObjectList.append({"name": sourceFile, "object": returnValue["object"]})
                break
            elif ERROR_FILE_NOT_FOUND == returnValue["error"]:
                sourceFileFullPath = getExcelFileName(f"Open {sourceFile}",sessionData["rootFolder"])
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
    #save the settings
    sessionData["sessionFileName"] = "session.json"
    if not sessionData["sessionFileName"] == None:
        #save the session.
        saveSessionData(sessionData["sessionFileName"], sessionData["pdfFileName"], sessionData["templateFileName"], fileObjectList)
    print ("Start Processing [", len(recordIDList) , "] records")
    for recordId in recordIDList:
        messageHolder = {"id": 0, "action": MESSAGE_CLEAR, "message": None}
        #update_message(messageHolder,MESSAGE_NEW,templateFileName,True)
        #windowName = "Status of Record ID = [" + str(recordId["identifier"]) + "]"
        outputFileName = str(recordId["key"])+"-"+str(recordId["identifier"])+".pdf"
        print("Processing ["+ outputFileName + "]")
        # Start a background thread to simulate message updates
        thread = threading.Thread(target=processRecord, args=(messageHolder,fileObjectList,recordId,textOverlayList,sessionData["pdfFileName"],PDF_FIRST_PAGE,outputFileName))
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


if __name__ == "__main__":
    main()


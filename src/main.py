import threading
import time
from constants.templatedata import TEMPLATE_SHEET_NAME, TEMPLATE_FOLDER_NAME, RECORD_LIST_SHEET_NAME
from constants.errorcodes import ERROR_FILE_NOT_FOUND, ERROR_SUCCESS, ERROR_FILE_ENCRYPTED, ERROR_UNKNOWN, ERROR_ITEM_NOT_FOUND
from constants.pdfData import PDF_FIRST_PAGE, getPdfPage
from projectutils.guifunc import WINDOW_QUIT # Import GUI constants, Window
from projectutils.guifunc import MESSAGE_NEW, MESSAGE_ADD, MESSAGE_CLEAR # Import GUI constants Message
from projectutils.guifunc import showStatus, getExcelFileName, getPassword, getPdfFileName  # Import GUI functions
from projectutils.businessfunc import loadTemplateData, getFilesFromOverlayList, loadRecordIdList
from projectutils.businessfunc import getStringFromFileObject, concatString
from projectutils.filefunc import openExcelFile, createTempFile, removeFiles
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
        #windowName = "Status of Record ID = [" + str(recordId["identifier"]) + "]"
        outputFileName = str(recordId["key"])+"-"+str(recordId["identifier"])+".pdf"
        # Start a background thread to simulate message updates
        thread = threading.Thread(target=processRecord, args=(messageHolder,fileObjectList,recordId,textOverlayList,pdfFileName,PDF_FIRST_PAGE,outputFileName))
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

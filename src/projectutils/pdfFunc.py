""" PDF Functions 
    src/projectutils/pdfFunc.py
    This file contains the functions that use PyPDF library.
    Author: vipulasrilanka@yahoo.com 
    (c) 2024 """

from PyPDF2 import PdfWriter, PdfReader
import io
import re #consider moving to a business function
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, legal, A4
from reportlab.pdfgen.textobject import PDFTextObject 
from constants.errorcodes import ERROR_SUCCESS, ERROR_UNKNOWN, ERROR_LONG_TEXT
from constants.pdfData import PDF_FIRST_PAGE, PDF_DEFAULT_LINE_SPACE, PDF_DEFAULT_LINE_SPACE_FACTOR, PDF_DEFAULT_FONT, PDF_DEFAULT_FONT_SIZE

def addOverlayToPdf(PdfTemplateName, PdfTemplatePage, outputFileName, pdfOverlayList):
    """ Create a new PDF from PdfTemplateName, with overlay text. Please note the output
        page is always a single page. <TODO> Add support for multipage documents
    Args:   PdfTemplateName (string) : Template PDF file name
            PdfTemplatePage (int) : Zero based template page number
            outputFileName (string) : output file name
            pdfOverlayList (list) : List of directories
    Returns: int: Error codes"""

    print("+Fn addOverlayToPdf", PdfTemplateName,outputFileName)
    #create a canvas and add the overlay data
    overlayByteIO = io.BytesIO()
    overlayCanvas = canvas.Canvas(overlayByteIO, pagesize=letter)
    #process multiline if any
    for overlay in pdfOverlayList :
        #mandatory fields
        print(overlay["param"]["X"], overlay["param"]["Y"], overlay["string"])
        overlay = setOptionalParams(overlay)
        #process the text Line
        textObj = getTextObj(overlayCanvas,overlay["string"],overlay["param"])
        if not isinstance(textObj, PDFTextObject):
            print("ERROR [- Fn addOverlayToPdf]: Can not find the text Object: Bad arguments")
            return ERROR_UNKNOWN
        else:
            overlayCanvas.drawText(textObj)
    overlayCanvas.save()
    #move to the beginning of the StringIO buffer
    overlayByteIO.seek(0)
    print("create a blank PDF with overlay")
    try:
        # create a new PDF with text overlay
        overlayPdf = PdfReader(overlayByteIO)
        # read your existing PDF
        print("Open PDF template")
        templatePdf = PdfReader(open(PdfTemplateName, "rb"))
        output = PdfWriter()
        # add the "watermark" (which is the new pdf) on the existing page
        page = templatePdf.pages[PdfTemplatePage]
        print("Merge the overlay now..!")
        page.merge_page(overlayPdf.pages[PDF_FIRST_PAGE]) #overlayPdf has only one page
        print("Add the page to ourputFile..")
        output.add_page(page)
        # finally, write "output" to a real file
        print("Open Output file")
        outPutFile = open(outputFileName, "wb")
        print("Write to File..")
        output.write(outPutFile)
        outPutFile.close()
    except Exception as e:
        print(f"Error: {e}")
        print("ERROR [- Fn addOverlayToPdf]")
        return ERROR_UNKNOWN
    print("-Fn addOverlayToPdf")
    return ERROR_SUCCESS


def setOptionalParams(overlay):
    """ Some of the parameters are optional in the template, but
        mandatory for the textObject. So set them to default"""
    if not "param" in overlay:
        return ERROR_UNKNOWN
    if not "Font" in overlay["param"]:
        overlay["param"]["Font"] = PDF_DEFAULT_FONT
    if not "FontSize" in overlay["param"]:
        overlay["param"]["FontSize"] = PDF_DEFAULT_FONT_SIZE
    if not "LineSpace" in overlay["param"]:
        overlay["param"]["LineSpace"] = PDF_DEFAULT_LINE_SPACE
    if not "Function" in overlay["param"]:
        overlay["param"]["Function"] = None

    return overlay

def getTextObj(canvas,text, params):
    """ returns a text object with the data given. The text object has the capability of 
        holding multiple lines with different formats."""
    print("Fn getTextLine")
    textLines = []
    lineSpace = getLineHeight(params["FontSize"], params["LineSpace"])
    if None == params["Function"]:
        print("[addOverlayToPdf] No extra text proccesisng")
        textLines.append({"text": (text), "fontSize": params["FontSize"], "lineSpace": lineSpace, "set_cursor": None})
    else:
        #Breaks the text in to lines and adjust font for each line based on rules
        print("[addOverlayToPdf] Call text proccesing")
        textLines = processFunc(canvas,text, params["Font"], params["FontSize"], params["Function"])
        if not isinstance(textLines,list):
            print ("ERROR [getTextObj]. Can not print emplty line")
            return ERROR_UNKNOWN
    textObj = canvas.beginText( params["X"],  params["Y"])
    #store text lines based on rules
    for textLine in textLines:
        if isinstance(textLine["fontSize"],int):
            textObj.setFont(params["Font"],textLine["fontSize"])
        if isinstance(textLine["lineSpace"],int):
            lineSpace = textLine["lineSpace"]
        if isinstance(textLine["set_cursor"],int):
            textObj.moveCursor(textLine["set_cursor"], lineSpace)
        else:
            print("[getTextObj] not moving curser.!")
        textObj.textOut(textLine["text"])
    return textObj


def getLineHeight(fontSize, lineSpace):
    """ Retuens the lince ppace based on given rules. Value is in points"""
    if not fontSize:
        return PDF_DEFAULT_LINE_SPACE
    if not lineSpace:
        lineSpace = fontSize * PDF_DEFAULT_LINE_SPACE_FACTOR
    elif isinstance(lineSpace,float) or isinstance(lineSpace,int):
        return lineSpace
    elif lineSpace.endswith('X'):
        return float(lineSpace[:-1]) * fontSize
    else:
        return fontSize * PDF_DEFAULT_LINE_SPACE_FACTOR
        

def processFunc(canvas, text, font, fontSize, function):
    """ Fuction to alter the text """
#    paramList = parseFunctionString(paramList)
    if not "SrinkToFit" == function["name"]:
        print("Error: [processFunc] not a supported Function ")
        return text
    width = int(function["param1"])
    maxLines = int(function["param2"])
    textLines = []
    # Function to get the width of the text
    def getTextWidth(text, fontSize):
        return canvas.stringWidth(text, font, fontSize)
        # Try reducing font size until the text fits
    while True:
        # Set the font to the current size
        canvas.setFont(font, fontSize)
        words = text.split(' ')
        set_cursor = None
        currentLine = ""

        # Loop through words to form lines that fit within the width
        for word in words:
            testLine = currentLine + (" " + word if currentLine else word)
            if getTextWidth(testLine, fontSize) <= width:
                currentLine = testLine
            else:
                textLines.append({"text": (currentLine), "fontSize": fontSize, "lineSpace": None, "set_cursor": set_cursor})
                #we need to return cursor to 0 from next line onwards
                set_cursor = 0
                currentLine = word

        # Append the last line, regardless of how many lines have been created
        if currentLine:
            textLines.append({"text": (currentLine), "fontSize": fontSize, "lineSpace": None, "set_cursor": set_cursor})

        # Check if the number of lines is less than or equal to the allowed number of lines
        if len(textLines) <= maxLines:
            break

        # Reduce the font size if the text still doesn't fit
        fontSize -= 1
        if fontSize < 6:  # Set a minimum font size limit
            print ("ERROR [constWidth]. The text line is too long to fit to [" + str(width) + "] pixels x [" + str(maxLines) + "] lines")
            return ERROR_LONG_TEXT
    return textLines
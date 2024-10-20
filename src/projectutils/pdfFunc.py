""" PDF Functions 
    src/projectutils/pdfFunc.py
    This file contains the functions that use PyPDF library.
    Author: vipulasrilanka@yahoo.com 
    (c) 2024 """

from PyPDF2 import PdfWriter, PdfReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, legal, A4
from constants.errorcodes import ERROR_SUCCESS, ERROR_UNKNOWN
from constants.pdfData import PDF_FIRST_PAGE, getPdfPage

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
    for overlay in pdfOverlayList :
        print(overlay["x"], overlay["y"], overlay["string"])
        overlayCanvas.drawString(overlay["x"], overlay["y"], overlay["string"])
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
        return ERROR_UNKNOWN
    print("-Fn addOverlayToPdf")
    return ERROR_SUCCESS


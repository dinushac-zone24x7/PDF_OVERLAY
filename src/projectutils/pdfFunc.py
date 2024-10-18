from PyPDF2 import PdfWriter, PdfReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, legal, A4
from constants.errorcodes import ERROR_SUCCESS, ERROR_UNKNOWN


def addOverlayToPdf(PdfTemplateName, outputFileName, pdfOverlayList):
    """Create a new PDF from PdfTemplateName, with overlay text"""
    print("+Fn addOverlayToPdf", PdfTemplateName,outputFileName)
    #create a canvas and add the overlay data
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    for overlay in pdfOverlayList :
        print(overlay["x"], overlay["y"], overlay["string"])
        can.drawString(overlay["x"], overlay["y"], overlay["string"])
    can.save()
    #move to the beginning of the StringIO buffer
    packet.seek(0)
    # create a new PDF with Reportlab
    print("create a blank PDF with overlay")
    try:
        blankPdf = PdfReader(packet)
        # read your existing PDF
        print("Open PDF template")
        templatePdf = PdfReader(open(PdfTemplateName, "rb"))
        output = PdfWriter()
        # add the "watermark" (which is the new pdf) on the existing page
        page = templatePdf.pages[0]
        print("Merge the overlay now..!")
        page.merge_page(blankPdf.pages[0])
        print("Add the page to ourputFile..")
        output.add_page(page)
        # finally, write "output" to a real file
        print("Open Output file")
        output_stream = open(outputFileName, "wb")
        print("Write to File..")
        output.write(output_stream)
        output_stream.close()
    except Exception as e:
        print(f"Error: {e}")
        return ERROR_UNKNOWN
    print("-Fn addOverlayToPdf")
    return ERROR_SUCCESS


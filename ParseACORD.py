# ParseACORD.py - Parser for ACORD PDF	
# Python 3

# Usage:
#   python ParseACORD.py <InputPDFFileNameWithoutExtension>
# Example:
#   python ParseACORD.py "ACORD_Life_Standards_PubDoc_2.38.00"
#   - uses input file ".\\ACORD_Life_Standards_PubDoc_2.38.00"

# Imports
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
import pdfminer  # pdfminer reads input PDF report
import sqlite3   # sqlite3 database stores all temporary data
#import re        # regular expressions are used to match and parse text strings
import sys       # for command line parameter retrieval
#from __future__ import print_function

def init_db(cur):

    cur.execute(open('.\\CREATE TABLE ParseACORDIn.sql', 'r').read())
        
def add_row(cur, pDFFileName, pageNum, leftPoints, lowerPoints, rightPoints, upperPoints, pDFString):
    
    # Convert position measures from points to ingtegral millipoints
    leftMP   = int(round(leftPoints*1000.0))
    lowerMP  = int(round(lowerPoints*1000.0))
    rightMP  = int(round(rightPoints*1000.0))
    upperMP  = int(round(upperPoints*1000.0))
    
    # Calculate width, height and midpoints in millipoints
    widthMP  = int(round((rightPoints-leftPoints)*1000.0))
    heightMP = int(round((upperPoints-lowerPoints)*1000.0))
    hMidPtMP = (leftPoints*1000.0 + rightPoints*1000.0) / 2
    vMidPtMP = (upperPoints*1000.0 + lowerPoints*1000.0) / 2
    
    # Insert parsed string into database
    cur.execute('''
       INSERT INTO ParseACORDIn (
           PDFFileName,
           PageNum, 
           LeftMP, 
           LowerMP, 
           RightMP, 
           UpperMP, 
           WidthMP, 
           HeightMP, 
           VCumulativeMP, 
           PDFString)
       VALUES (?,?,?,?,?,?,?,?,?,?)''', (
           pDFFileName,
           pageNum, 
           leftMP,
           lowerMP,
           rightMP,
           upperMP,
           widthMP,
           heightMP,
           pageNum*11*72*1000-int(round(vMidPtMP)),
           pDFString[0:-1] ) )

def parse_obj(cur, pDFFileName, lt_objs):

    # loop over the object list
    for obj in lt_objs:

        # if it's a textbox, record text and location
        if isinstance(obj, pdfminer.layout.LTTextLineHorizontal):
            # if  int(round(obj.bbox[1]*1000.0)) > 22140:
            add_row(cur, pDFFileName, myPageNum, obj.bbox[0], obj.bbox[1], obj.bbox[2], obj.bbox[3], obj.get_text())
            #f.write(obj.get_text().replace(u"\u25cf", "* ").replace(u"\u25a0", "* ").replace(u"\u25b2", "*3 ").replace(u"\u25bc", "*4 ").replace(u"\u2192", "->"))
            #f.write(u"\n")

        # if it's a container, recurse
        elif isinstance(obj, pdfminer.layout.LTFigure):
            parse_obj(cur, pDFFileName, obj._objs)

        # if it's a container, recurse
        elif isinstance(obj, pdfminer.layout.LTTextBox):
            # if  int(round(obj.bbox[1]*1000.0)) > 22140:
            add_row(cur, pDFFileName, myPageNum, obj.bbox[0], obj.bbox[1], obj.bbox[2], obj.bbox[3], obj.get_text())
            #f.write(obj.get_text().replace(u"\u25cf", "*1 ").replace(u"\u25a0", "*2 ").replace(u"\u25b2", "*3 ").replace(u"\u25bc", "*4 ").replace(u"\u2192", "->"))




########################
# Program starts here: #
########################

# Connect to a new sqlite database
db = sqlite3.connect(".\\" + str(sys.argv[1]) + ".sqlite")
cur = db.cursor()
init_db(cur)

# Open the input PDF report
fp = open(".\\" + str(sys.argv[1]) + ".pdf", 'rb')

# Create a PDF parser object associated with the file object.
parser = PDFParser(fp)

# Create a PDF document object that stores the document structure.
# Password for initialization as 2nd parameter
document = PDFDocument(parser)

# Check if the document allows text extraction. If not, abort.
#if not document.is_extractable:
#    raise PDFTextExtractionNotAllowed

# Create a PDF resource manager object that stores shared resources.
rsrcmgr = PDFResourceManager()

# Create a PDF device object.
device = PDFDevice(rsrcmgr)

# BEGIN LAYOUT ANALYSIS
# Set parameters for analysis.
laparams = LAParams()

# Create a PDF page aggregator object.
device = PDFPageAggregator(rsrcmgr, laparams=laparams)

# Create a PDF interpreter object.
interpreter = PDFPageInterpreter(rsrcmgr, device)

# initialize the page number
myPageNum = 1

# 
#f = open(".\\Parsed " + str(sys.argv[1]) + ".txt", 'w')

# loop over all pages in the document
for page in PDFPage.create_pages(document):
    
    # read the page into a layout object
    interpreter.process_page(page)
    layout = device.get_result()
    
    # extract text from this object
    parse_obj(cur, str(sys.argv[1]), layout._objs)
    
    # commit the new page's information to the database
    db.commit()
    
    # advance the page number
    myPageNum = myPageNum + 1

db.commit()
db.close()
#f.close()
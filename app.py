from spire.xls import *
from spire.xls.common import *
import os


# Change based on your requirement
inputFile = "/var/myapp/index.xlsx"
outputFile = "/var/www/html/index.html"


while True:
    try:
        # Create a Workbook instance
        workbook = Workbook()

        # Load a sample Excel file
        workbook.LoadFromFile(inputFile)

        # Get the first sheet of this file
        sheet = workbook.Worksheets[0]

        # Save the worksheet to HTML
        sheet.SaveToHtml(outputFile)
        workbook.Dispose()

        os.remove("sample_1.xlsx")

    except:
        continue
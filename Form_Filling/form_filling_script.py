import csv
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import comtypes.client

# Load the .docx template
doc = DocxTemplate("template.docx")

# Open the .csv file
with open("data.csv", "r") as csvfile:
    reader = csv.DictReader(csvfile)

    # Iterate through each row in the .csv file
    for row in reader:
        # Create a context dictionary with the data from the .csv row
        context = row

        # Render the .docx template with the data from the .csv row
        doc.render(context)

        # Save the filled form as a .docx file
        doc.save(f"{row['name']}.docx")

        # Convert the .docx file to a .pdf file
        word = comtypes.client.CreateObject("Word.Application")
        doc = word.Documents.Open(f"{row['name']}.docx")
        doc.SaveAs(f"{row['name']}.pdf", FileFormat=17)
        doc.Close()

    word.Quit()
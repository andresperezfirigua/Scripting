import csv
from docxtpl import DocxTemplate
import datetime

# Load the .docx template
doc = DocxTemplate("Asignation_Template.docx")

# Open the .csv file
with open("data.csv", "r") as csvfile:
    reader = csv.DictReader(csvfile, delimiter=';')

    # Iterate through each row in the .csv file
    for row in reader:
        # Create a context dictionary with the data from the .csv row
        context = row

        # Edit LOB from data
        context['lob'] = context['lob'].split(' - ')[1]

        # Edit serial number
        context['serial_number'] = f'00{context["serial_number"]}'

        # Add current date
        date = datetime.datetime.now()
        context['date'] = date.strftime('%B %d, %Y')

        # Render the .docx template with the data from the .csv row
        doc.render(context, autoescape=True)

        # Save the filled form as a .docx file
        file_name = f"{row['emp_eeid']} - {row['emp_name']}"
        doc.save(f"Filled_Forms/Linda_Rozo/{file_name}.docx")

        print(f"Filled_Forms/Linda_Rozo/{file_name}.docx")

    # word.Quit()

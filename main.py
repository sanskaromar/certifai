import os
from pptx import Presentation
import comtypes.client


def PPT_to_PDF(input_pptx, output_pdf, formatType=32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if output_pdf[-3:] != "pdf":
        output_pdf = output_pdf + ".pdf"
    deck = powerpoint.Presentations.Open(input_pptx)
    deck.SaveAs(output_pdf, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


# Load the existing PowerPoint presentation
input_pptx = "existing_presentation.pptx"
output_pdf = "output_presentation.pdf"

base_path = os.path.dirname(os.path.abspath(__file__))

prs = Presentation(input_pptx)

# Define the placeholder text to replace
placeholder_name = "Placeholder_Name"

# Get user input for the replacement name
replacement_name = input("Enter the name to replace '{}': ".format(placeholder_name))

# Iterate through slides and shapes to find and replace the placeholder text
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if placeholder_name in run.text:
                        run.text = run.text.replace(placeholder_name, replacement_name)

# Save the modified PowerPoint presentation
updated_pptx = "updated_presentation.pptx"
prs.save(updated_pptx)

# Convert the updated PowerPoint presentation to PDF
PPT_to_PDF(os.path.join(base_path, updated_pptx), os.path.join(base_path, output_pdf))

print(f"The presentation has been updated and saved as '{updated_pptx}'.")

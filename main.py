from pptx import Presentation

# Load the existing PowerPoint presentation
input_pptx = 'existing_presentation.pptx'
output_pdf = 'output_presentation.pdf'

prs = Presentation(input_pptx)

# Define the placeholder text to replace
placeholder_name = 'Placeholder_Name'

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
updated_pptx = 'updated_presentation.pptx'
prs.save(updated_pptx)

print(f"The presentation has been updated and saved as '{updated_pptx}'.")

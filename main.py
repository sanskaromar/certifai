import os
import csv
import time
from pptx import Presentation
import comtypes.client


def PPT_to_PDF(input_pptx, output_pdf, formatType=32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if output_pdf[-3:] != "pdf":
        output_pdf = output_pdf + ".pdf"
    deck = powerpoint.Presentations.Open(input_pptx, WithWindow=False)
    deck.SaveAs(output_pdf, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


def process_pptx(row, save_pptx=False):
    # Load the existing PowerPoint presentation
    input_pptx = "certificate_template.pptx"

    base_path = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(base_path, "certificates_pdf")

    prs = Presentation(input_pptx)

    # Define the placeholders to replace
    placeholders = {
        "Placeholder_Name": row["name"],
        "Placeholder_refno": row["id"],
    }

    # Iterate through slides and shapes to find and replace the placeholders
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for placeholder, value in placeholders.items():
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, value)

    # Save the modified PowerPoint presentation
    updated_pptx = f"{row['id']}_presentation.pptx"

    if save_pptx:
        output_folder = os.path.join(base_path, "certificates_pptx")
        output_pptx = os.path.join(output_folder, updated_pptx)
        prs.save(output_pptx)

    # Convert the updated PowerPoint presentation to PDF
    output_pdf = os.path.join(output_folder, f"{row['id']}.pdf")
    PPT_to_PDF(updated_pptx, output_pdf)

    print(f"Certificate for {row['name']} has been saved as '{output_pdf}'.")


def main():
    # Load data from CSV file
    csv_file = "data.csv"
    start_time = time.time()
    with open(csv_file, mode="r") as file:
        reader = csv.DictReader(file)
        for row in reader:
            print("Generating certificate for {}...".format(row["name"]))
            # Generate certificate for each row
            process_pptx(row)

    time_elapsed = time.time() - start_time
    print(
        "Time elapsed for generating {} certificates: {:.2f} seconds.".format(
            reader.line_num - 1, time_elapsed
        )
    )
    print("All certificates have been generated.")


if __name__ == "__main__":
    main()

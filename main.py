import os
import csv
import time
import pandas as pd
from pptx import Presentation
import comtypes.client

pd.options.mode.chained_assignment = None  # default='warn'


def PPT_to_PDF(input_pptx, output_pdf, formatType=32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if output_pdf[-3:] != "pdf":
        output_pdf = output_pdf + ".pdf"
    deck = powerpoint.Presentations.Open(input_pptx, WithWindow=False)
    deck.SaveAs(output_pdf, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


def process_pptx(row):
    # Load the existing PowerPoint presentation
    input_pptx = "certificate_template.pptx"

    base_path = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(base_path, "certificates_pdf")
    pptx_output_folder = os.path.join(base_path, "certificates_pptx")
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(pptx_output_folder, exist_ok=True)

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
    updated_pptx = os.path.join(pptx_output_folder, f"{row['id']}.pptx")
    prs.save(updated_pptx)

    # Convert the updated PowerPoint presentation to PDF
    output_pdf = os.path.join(output_folder, f"{row['id']}.pdf")
    PPT_to_PDF(updated_pptx, output_pdf)

    print(f"Certificate for {row['name']} has been saved as '{output_pdf}'.")


def generate_certificates_summary(df):
    """
    Generate a Markdown file with a summary of the certificates generated.
    Can be used to validate the certificates generated.
    """
    certificates_md = "certificates.md"
    # Create a DataFrame for the Markdown table
    df_md = df[["id", "name", "profile_url"]]
    df_md["profile_url"] = df_md["profile_url"].apply(
        lambda x: f"[Link]({x})" if pd.notna(x) else ""
    )

    # Write the DataFrame to the Markdown file
    with open(certificates_md, mode="w", newline="") as md_file:
        md_file.write("# Certificates\n")
        time_now = time.strftime("%d %B %Y, %H:%M:%S", time.localtime())
        md_file.write(f"Generated on {time_now}\n\n")
        df_md.to_markdown(md_file, index=False)


def main():
    start_time = time.time()
    # Load data from CSV file
    csv_file = "data.csv"
    df = pd.read_csv(csv_file)

    for _, row in df.iterrows():
        print("Generating certificate for {}...".format(row["name"]))
        process_pptx(row)

    generate_certificates_summary(df)

    time_elapsed = time.time() - start_time
    print(
        "Time elapsed for generating {} certificates: {:.2f} seconds.".format(
            len(df), time_elapsed
        )
    )
    print("All certificates have been generated.")


if __name__ == "__main__":
    main()

# Certifai: Certificate Generator
[![Leave a Star](https://img.shields.io/github/stars/sanskaromar/certifai?style=social)](https://github.com/sanskaromar/certifai)
## Description
Certifai is a highly customizable Certificate Generator built using Python, designed to streamline the certificate generation process. It automates the creation of certificates in PDF format using a PowerPoint template with predefined placeholders. The script pulls data from a CSV file to populate these placeholders, adds QR codes to the certificates, and generates a summary in Markdown format. One of its standout features is the use of multiprocessing, allowing you to harness the power of parallel processing for faster execution.

This README will provide you with detailed instructions on setting up and running the script, giving you full control over your certificate generation process.

### Certificate Verification

Certificates generated using this generator are designed with authenticity in mind. Each certificate contains a reference number, and you can easily verify their authenticity by comparing these reference numbers with the ones listed in the `certificates.md` file. This `certificates.md` file can be hosted anywhere, such as on a website or in a GitHub repository. For example, you can find a real-world example of this in action at [GDSC MNNIT's GitHub repository](https://github.com/gdsc-mnnita/Google-Cloud-Study-Jams/blob/main/GCSJ-23/README.md).

On each certificate, you will find the following information:

- **Certificate Reference Number**: This unique reference number helps you verify the certificate's authenticity.
- **Link to Repository**: A link to the repository where `certificates.md` is hosted for reference.
- **QR Code**: The certificate also includes a QR code that links to the candidate's profile URL, where badges and additional information can be viewed.

We have taken care to minimize the public exposure of candidate information to respect their privacy.

This additional information ensures that the generated certificates are not only beautifully designed but also secure and verifiable.

## Prerequisites

- Python 3.7 or higher
- A virtual environment (optional but recommended)
- Powerpoint installed on your windows OS

## Setup

1. Clone this repository to your local machine or download the project files.

2. Create a virtual environment (optional but recommended):
   ```bash
   python -m venv venv
   venv\Scripts\Activate.ps1
   ```

3. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

4. Prepare your configuration:
   - Create a `config.json` file with the following structure:
     ```json
     {
         "input_pptx": "certificate_template.pptx",
         "output_folder": "certificates_pdf",
         "logo_path": "logo.png",
         "csv_file": "data.csv"
     }
     ```
   - Replace the values with your own file paths.
   - `input_pptx`: The path to the PowerPoint template.
   - `output_folder`: The directory where PDF certificates will be saved.
   - `logo_path`: The path to the logo image.
   - `csv_file`: The path to the CSV file containing data.

5. Prepare your data:
   - Create a `data.csv` file with a header row and data in the following format:
     ```
     email, certificate_id,name,profile_url
     john@example.com,GCSJ23C2-MNNITAx078,John Doe,https://example.com/johndoe
     jane@example.com,GCSJ23C2-MNNITAx078,Jane Smith,https://example.com/janesmith
     ...
     ```
     NOTE: The first column email is not necessary an you can have any other column or any constant value in the first column. The initial idea was to mail the certificates to the participants but that is not implemented yet.

## Project Structure Explained

The Certificate Generator project consists of the following files and folders:

- `README.md`: The main documentation you are currently reading.
- `LICENSE.md`: The project's licensing information.
- `.gitignore`: Specifies which files or directories should be ignored by Git.
- `requirements.txt`: Contains a list of Python dependencies for the project.
- `config.json`: Configuration file that allows users to customize the script.
- `certificate_generator.py`: The main script for generating certificates.
- `data.csv`: Data file that includes participant information.
- `logo.png`: The logo of your organization, which is included in the QR codes.
- `certificate_template.pptx`: The PowerPoint template used for generating certificates.

**Note:** Ensure that your certificate template includes text with placeholders for the participant's name and certificate number, labeled as `Placeholder_Name` and `Placeholder_refno`, respectively. You may also need to adjust the QR code's position based on your template.

### Output Files and Folders

The generated certificates and summary will be saved in the following locations:

- `certificates_pdf/`: The directory containing the generated PDF certificates.
- `certificates.md`: A Markdown summary file listing the generated certificates.

### Temporary Generated Folders

Temporary folders are created during script execution and are deleted when the program is closed. (You don't have to worry about them.)

- `certificates_pptx/`: Temporary directory for generated PowerPoint certificates.
- `qr_codes/`: Temporary directory for generated QR codes.


## Usage

1. Run the Certificate Generator:
   ```bash
   python main.py
   ```

2. The script will populate the placeholders in the PowerPoint template with the data from the CSV file, generate QR codes for each certificate, and save PDF certificate files in the given output folder.

3. A Markdown summary file (`certificates.md`) will be created in the project directory, which you can use to validate the certificates generated.

## Customization

- You can modify the `config.json` file to change the file paths and settings.
- Customize the PowerPoint template with your own design and placeholders.
- Adjust the QR code generation options and image settings in the code if needed.

## Advanced Configuration

For more advanced configuration options, you can modify the Python script directly. It provides functions for generating QR codes, replacing placeholders, and processing PowerPoint files. You can further customize the script to fit your specific requirements.

## License

This project is licensed under the following License - see the [LICENSE](LICENSE.md) file for details.

TLDR; If you use this project, you must give credits to the author by linking to this repository.

## Performance
This Certificate Generator is a powerful tool that efficiently generates certificates from a PowerPoint template. The generation speed is dependent upon your system specs, but it ulitizes multiprocessing to maximize the performance.

For my actual use case, while generating certificates for [Google Cloud Study Jams participants from MNNIT](https://github.com/gdsc-mnnita/Google-Cloud-Study-Jams/tree/main/GCSJ-23), it created 103 certificates in just 31 seconds on average.

# Limitations

Currently, I have only implemented it for Windows OS and utilized PowerPoint for pptx to PDF conversion. It is possible to implement it in Linux, but that is not done yet. Also, since the tasks here are mostly I/O bound, multithreading may result in even better performance, but that is not implemented as well.

# Example Usage

Please refer to the GDSC MNNIT's [Google Cloud Study Jams repository](https://github.com/gdsc-mnnita/Google-Cloud-Study-Jams/tree/main/GCSJ-23) for an example of how I used this Certificate Generator to generate certificates for the participants of the event.

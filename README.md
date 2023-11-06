# Word to PDF Converter

A convenient tool for converting all Word documents in a specified folder to PDF format and organizing them neatly.

## Features

- Batch converts `.doc` and `.docx` files to PDF.
- Saves converted PDFs in a designated output folder.

## Requirements

- Python 3.x
- Microsoft Word must be installed on your system.
- `comtypes` Python module (install with `pip install comtypes`).

## Setup

Before running the script, ensure you update the following variables in the `word_to_pdf_converter.py` file:

- `SOURCE_DIRECTORY`: The path to the folder containing your Word documents.
- `DESTINATION_DIRECTORY`: The path to the folder where you want the PDFs to be saved.

## Usage

After configuring the script, run it with:

```bash
python word_to_pdf_converter.py

The script will find all Word documents in the SOURCE_DIRECTORY and convert them to PDF files in the DESTINATION_DIRECTORY.
Note

This script requires Microsoft Word to be installed on your system as it uses the Word application for conversion.
Disclaimer

This script is provided for educational and professional purposes. Please ensure that you have the right to convert and handle the documents before use. The author is not responsible for any misuse or damage that might occur.

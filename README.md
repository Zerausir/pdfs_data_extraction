# PDF Extractor - README

## Overview

This project provides a Python script to process PDF files, extract structured information using regular expressions,
and export the data to an Excel file with filters and styled formatting.

## Features

- Extracts key information from PDFs, including document name, date, subject, references, and annexes.
- Processes all PDFs in a specified directory.
- Exports the extracted data into an Excel file with formatted headers and adjustable column widths.
- Validates required environment variables for seamless configuration.

## Requirements

- Python 3.8 or later
- Dependencies (see `requirements.txt`):
    - pandas
    - PyPDF2
    - openpyxl
    - environs

## Installation

1. Clone the repository.
2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Configuration

Ensure that the following environment variables are set in a `.env` file in the root directory:

```env
SERVER_ROUTE=path/to/pdf/directory
DOWNLOAD_ROUTE=path/to/save/excel
```

- `SERVER_ROUTE`: Directory containing the PDF files to process.
- `DOWNLOAD_ROUTE`: Directory where the output Excel file will be saved.

## Usage

1. Run the script:
   ```bash
   python script_name.py
   ```
   Replace `script_name.py` with the name of the script file.

2. The script will process all PDFs in the specified directory and save the extracted data in an Excel file at the
   output location defined in `DOWNLOAD_ROUTE`.

## Code Structure

### `PDFExtractor` Class

Handles PDF processing:

- **`extract_text_from_pdf(pdf_path)`**: Extracts the full text from a PDF file.
- **`extract_field_with_regex(text, pattern)`**: Extracts a specific field using a regex pattern.
- **`process_pdf(pdf_path)`**: Extracts all required fields from a single PDF.
- **`process_directory(directory)`**: Processes all PDFs in a directory and returns a consolidated DataFrame.

### Utility Functions

- **`save_to_excel_with_style(df, output_file)`**: Exports a DataFrame to an Excel file with styles, filters, and
  adjustable column widths.
- **`verify_environment_variables()`**: Validates that all required environment variables are set.

## Error Handling

- Catches exceptions during PDF processing and logs the errors.
- Validates environment variables and raises an error if any are missing.

## Example Output

An Excel file with the following columns:

- `Nombre`: Name of the document.
- `Fecha`: Date extracted from the document.
- `Asunto`: Subject of the document.
- `Anexo`: List of annexes.
- `Referencias`: References mentioned in the document.

## Contributing

1. Fork the repository.
2. Create a feature branch:
   ```bash
   git checkout -b feature-name
   ```
3. Commit your changes:
   ```bash
   git commit -m "Add new feature"
   ```
4. Push to the branch:
   ```bash
   git push origin feature-name
   ```
5. Open a pull request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Acknowledgments

- Libraries
  used: [PyPDF2](https://pypi.org/project/PyPDF2/), [pandas](https://pandas.pydata.org/), [openpyxl](https://openpyxl.readthedocs.io/), [environs](https://pypi.org/project/environs/).

## Contact

For questions or feedback, please contact [Iván Suárez](https://github.com/Zerausir)

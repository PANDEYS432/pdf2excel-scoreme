# pdf2excel-scoreme

**pdf2excel-scoreme** is a Python application designed to extract tabular data from PDF files and convert it into Excel spreadsheets. This tool supports both text-based and scanned PDFs, utilizing PyMuPDF for text extraction and Tesseract OCR for processing scanned documents.


![Screenshot (90)](https://github.com/user-attachments/assets/42982e46-ad2d-45ff-99a1-a530d09b4405)
![Screenshot (91)](https://github.com/user-attachments/assets/885d5354-2a76-4c1c-8fda-da4b73f9bae7)


- Extracts tables from text-based PDFs.
- Performs OCR on scanned PDFs to extract text.
- Saves extracted data into Excel files for easy access and analysis.
- Automatically detects whether a PDF is text-based or scanned.
-  Handle tables with borders and without borders.

## Requirements

- Python 3.10.10
- PyMuPDF
- Tesseract OCR
- pytesseract
- Pillow
- pandas
- openpyxl
- numpy

## Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/PANDEYS432/pdf2excel-scoreme
   cd pdf2excel-scoreme
2. Create a virtual environment
   
   ```bash
   python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
3. Install the required packages:
   
   ```bash
   pip install -r requirements.txt
 ```
4. Install Tesseract OCR and set path in system environment

 ## Usage

1. Place your PDF files in the `input_pdfs` directory.

2. Run the script:

   ```bash
   python pdf_table_extractor.py
3. The extracted tables will be saved in the output_excels directory as Excel files.

## Contributing

Contributions are welcome! If you have suggestions for improvements or new features, please open an issue or submit a pull request.

# Invoice OCR Batch Processing
This Python project automates the extraction of key data from invoice images using Tesseract OCR and OpenCV. It is designed to handle batch processing of multiple invoices stored in a folder, extracting essential fields like Invoice Number, Date, and Total Amount, validating completeness, and saving all results into a structured Excel file.
The script supports English and Arabic invoices, performs basic image preprocessing, and provides a summary of missing fields for incomplete invoices. It is ideal for businesses or developers who need a quick and reliable way to digitize invoices for record-keeping, reporting, or integration with other systems.

---

## Features
- Batch process multiple invoice images from a folder  
- Extract essential invoice data:
  - Invoice Number  
  - Date  
  - Total Amount  
- Validate required fields and flag incomplete invoices  
- Support for **English and Arabic** invoices  
- Image preprocessing to improve OCR accuracy (grayscale, Gaussian blur, thresholding)  
- Append results to existing Excel file or create a new one  
- Track processing timestamp and source file name  

---

## Requirements
- Python 3.8+  
- Libraries:
```bash
pip install opencv-python pytesseract pandas openpyxl
```
- Tesseract OCR installed and path configured:
  ```
  pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"
  ```
## Usage
- Place all invoice images (PNG, JPG, JPEG, TIFF) in a folder, e.g., images/.
- Set the ```INPUT_FOLDER``` and ```OUTPUT_EXCEL``` paths in the script.
- After processing, check the ```data/invoices.xlsx``` file for extracted data.

## Example Output
| Invoice Number | Date	| Total Amount | File Name | Processed At | Status | Missing Fields |  
|----------------|------|--------------|-----------|--------------|--------|----------------|
| INV-001 | 2026-01-31 | 450.00 | invoice1.png | 2026-02-03 13:00:00 | Completed | None |    
| INV-002 | | 780.00 | invoice2.png | 2026-02-03 13:01:00 | Incomplete | Date |  


## How It Works
- The script loads each image and converts it to grayscale.
- Applies Gaussian blur and thresholding to enhance text for OCR.
- Uses Tesseract OCR to extract text from the image.
- Validates if the document contains the word invoice (English or Arabic).
- Extracts ```Invoice Number```, ```Date```, and ```Total Amount``` using regex.
- Marks missing fields and sets invoice status as ```Completed``` or ```Incomplete```.
- Saves all results to an Excel spreadsheet, appending new invoices to existing records.

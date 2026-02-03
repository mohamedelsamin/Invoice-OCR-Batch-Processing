import cv2
import pytesseract
import pandas as pd
import re
import os
from datetime import datetime

# CONFIGURATION
pytesseract.pytesseract.tesseract_cmd = r"D:\tesseract\tesseract.exe"

INPUT_FOLDER = "images"           # Folder containing invoice images
OUTPUT_EXCEL = "data/invoices.xlsx"
REQUIRED_FIELDS = ["Invoice Number", "Date", "Total Amount"]

# HELPER FUNCTION: Process single invoice image
def process_invoice(image_path):
    image = cv2.imread(image_path)
    if image is None:
        print(f"❌ Image not found: {image_path}")
        return None

    # Preprocessing
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

    # OCR
    text = pytesseract.image_to_string(thresh)
    if not text.strip():
        print(f"❌ OCR failed: {image_path}")
        return None

    # Document validation
    is_invoice = any(keyword in text.lower() for keyword in ["invoice", "فاتورة"])
    if not is_invoice:
        print(f"⚠️ Not detected as invoice: {image_path}")

    # Data extraction
    invoice_number = re.search(r"(Invoice\s*No|Invoice\s*#|فاتورة\s*رقم)[:\-]?\s*([\w\-]+)", text, re.IGNORECASE)
    date = re.search(r"(Date|التاريخ)[:\-]?\s*([\d/.-]+)", text)
    total = re.search(r"(Total|Grand\s*Total|الإجمالي)[:\-]?\s*\$?\s*([\d,.]+)", text)

    data = {
        "Invoice Number": invoice_number.group(2) if invoice_number else None,
        "Date": date.group(2) if date else None,
        "Total Amount": total.group(2) if total else None,
        "File Name": os.path.basename(image_path),
        "Processed At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # Validation
    missing_fields = [field for field in REQUIRED_FIELDS if not data[field]]
    data["Status"] = "Completed" if not missing_fields else "Incomplete"
    data["Missing Fields"] = ", ".join(missing_fields) if missing_fields else "None"

    return data

# MAIN: Process all images in folder
os.makedirs(os.path.dirname(OUTPUT_EXCEL), exist_ok=True)
all_data = []

for filename in os.listdir(INPUT_FOLDER):
    if filename.lower().endswith((".png", ".jpg", ".jpeg", ".tiff")):
        file_path = os.path.join(INPUT_FOLDER, filename)
        result = process_invoice(file_path)
        if result:
            all_data.append(result)

# SAVE TO EXCEL
if all_data:
    df_new = pd.DataFrame(all_data)
    if os.path.exists(OUTPUT_EXCEL):
        df_existing = pd.read_excel(OUTPUT_EXCEL, engine="openpyxl")
        df_final = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_final = df_new

    df_final.to_excel(OUTPUT_EXCEL, index=False, engine="openpyxl")
    print(f"✅ Batch processing completed. {len(all_data)} invoices processed.")
else:
    print("⚠️ No invoices were processed.")

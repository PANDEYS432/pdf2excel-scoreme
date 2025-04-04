import os
import fitz  
import pytesseract
from PIL import Image
from io import BytesIO
import pandas as pd
import re
import numpy as np
from collections import defaultdict

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

INPUT_DIR = "input_pdfs"
OUTPUT_DIR = "output_excels"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def is_scanned_pdf(doc, threshold=10):
    for page_num in range(min(3, len(doc))):
        page = doc[page_num]
        text = page.get_text().strip()
        if len(text) > threshold:
            return False
    return True

def preprocess_image(img):
    img = img.convert('L')
    threshold = 200
    img = img.point(lambda x: 0 if x < threshold else 255)
    return img

def ocr_pdf(doc):
    pages_data = []
    for i, page in enumerate(doc):
        print(f" OCR processing Page {i + 1}")
        pix = page.get_pixmap(dpi=300)
        img = Image.open(BytesIO(pix.tobytes("png")))
        img = preprocess_image(img)
        config = '--psm 6 --oem 3'
        text = pytesseract.image_to_string(img, config=config)
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        rows = defaultdict(list)
        y_tolerance = 10
        for i in range(len(data['text'])):
            if data['text'][i].strip():
                y_key = data['top'][i] // y_tolerance
                rows[y_key].append((data['left'][i], data['text'][i]))
        structured_data = []
        for y in sorted(rows.keys()):
            sorted_row = sorted(rows[y], key=lambda x: x[0])
            row_text = [text for _, text in sorted_row]
            if row_text:
                structured_data.append(row_text)
        pages_data.append(structured_data)
    return pages_data

def extract_text_based_tables(doc):
    pages_data = []
    for page_num, page in enumerate(doc):
        print(f" Parsing Page {page_num + 1}")
        blocks = page.get_text("dict")["blocks"]
        lines_with_coords = []
        for block in blocks:
            if "lines" not in block:
                continue
            for line in block["lines"]:
                spans = []
                line_bbox = line["bbox"]
                for span in line["spans"]:
                    spans.append({
                        "text": span["text"],
                        "bbox": span["bbox"],
                    })
                if spans:
                    lines_with_coords.append({
                        "y0": line_bbox[1],
                        "y1": line_bbox[3],
                        "spans": spans
                    })
        rows = group_lines_into_rows(lines_with_coords)
        table_data = []
        for row in rows:
            all_spans = []
            for line in row:
                all_spans.extend(line["spans"])
            all_spans.sort(key=lambda span: span["bbox"][0])
            row_text = [span["text"] for span in all_spans]
            if row_text:
                columns = analyze_and_split_row(row_text)
                table_data.append(columns)
        pages_data.append(table_data)
    return pages_data

def group_lines_into_rows(lines):
    if not lines:
        return []
    lines.sort(key=lambda line: line["y0"])
    rows = []
    current_row = [lines[0]]
    y_tolerance = 5
    for i in range(1, len(lines)):
        current_line = lines[i]
        previous_line = lines[i-1]
        if abs(current_line["y0"] - previous_line["y0"]) <= y_tolerance:
            current_row.append(current_line)
        else:
            rows.append(current_row)
            current_row = [current_line]
    if current_row:
        rows.append(current_row)
    return rows

def analyze_and_split_row(row_text):
    if len(row_text) > 1:
        return row_text
    text = ' '.join(row_text)
    columns = re.split(r'\s{2,}', text)
    if len(columns) > 1:
        return [col.strip() for col in columns if col.strip()]
    columns = re.split(r'\||\t', text)
    if len(columns) > 1:
        return [col.strip() for col in columns if col.strip()]
    numeric_pattern = r'\b\d+(?:[.,]\d+)?\b'
    matches = list(re.finditer(numeric_pattern, text))
    if len(matches) > 1:
        last_end = 0
        columns = []
        for match in matches:
            if match.start() - last_end > 3:
                if text[last_end:match.start()].strip():
                    columns.append(text[last_end:match.start()].strip())
            columns.append(match.group())
            last_end = match.end()
        if text[last_end:].strip():
            columns.append(text[last_end:].strip())
        if len(columns) > 1:
            return columns
    return [text]

def save_to_excel(pages_data, output_file):
    has_data = False
    for rows in pages_data:
        if rows and any(any(cell for cell in row) for row in rows):
            has_data = True
            break
    if not has_data:
        df = pd.DataFrame([["No table data detected in this PDF"]])
        df.to_excel(output_file, index=False, header=False)
        return
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        sheet_count = 0
        for i, rows in enumerate(pages_data):
            if not rows:
                continue
            cleaned_rows = []
            for row in rows:
                cleaned_row = [cell.strip() for cell in row]
                if any(cell for cell in cleaned_row):
                    cleaned_rows.append(cleaned_row)
            if cleaned_rows:
                sheet_count += 1
                sheet_name = f"Page_{i+1}"
                try:
                    df = pd.DataFrame(cleaned_rows)
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                except Exception as e:
                    print(f"Warning: Error with sheet {sheet_name}, writing without header: {str(e)}")
                    pd.DataFrame(cleaned_rows).to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        if sheet_count == 0:
            pd.DataFrame([["No table data detected in this PDF"]]).to_excel(
                writer, sheet_name="No_Data", index=False, header=False
            )

def main():
    files = [f for f in os.listdir(INPUT_DIR) if f.endswith(".pdf")]
    if len(files) < 2:
        print(" Not enough PDF files found in 'input_pdfs/'. Please ensure there are at least two PDF files.")
        return
    pdf_file_1 = os.path.join(INPUT_DIR, files[0])
    pdf_file_2 = os.path.join(INPUT_DIR, files[1])
    pdf_files = [pdf_file_1, pdf_file_2]
    
    for pdf_file in pdf_files:
        filename = os.path.basename(pdf_file)
        print(f"\n Processing: {filename}")
        try:
            doc = fitz.open(pdf_file)
            if pdf_file == pdf_file_2:
                print(" Detected scanned PDF - using OCR")
                pages_data = ocr_pdf(doc)
            else:
                print(" Detected text-based PDF - using direct extraction")
                pages_data = extract_text_based_tables(doc)
            output_filename = os.path.splitext(filename)[0] + ".xlsx"
            output_path = os.path.join(OUTPUT_DIR, output_filename)
            save_to_excel(pages_data, output_path)
            print(f" Saved Excel: {output_filename}")
        except Exception as e:
            print(f" Error processing {filename}: {str(e)}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    main()

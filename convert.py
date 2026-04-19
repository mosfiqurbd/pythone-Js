import utils.session
import os
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from docx import Document
from openpyxl import Workbook

#  SETTINGS 
PDF_PATH = r"C:\pdf_project\book.pdf"
OUTPUT_FOLDER = r"C:\pdf_project\images"
POPPLER_PATH = r"C:\poppler\Library\bin\poppler-25.12.0\Library\bin"
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# Step 1: PDF to Image & Rename
print("Step 1: PDF to Image converting...")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
pages = convert_from_path(PDF_PATH, dpi=300, first_page=1, last_page=20, poppler_path=POPPLER_PATH)
image_paths = []
for i, page in enumerate(pages):
    img_path = os.path.join(OUTPUT_FOLDER, f"page_{i+1:02d}.png")
    page.save(img_path, "PNG")
    image_paths.append(img_path)
    print(f"  ✅ page_{i+1:02d}.png saved")

# Step 2: Remove header/footer & OCR
print("\nStep 2: OCR [Image to text]...")
all_text = []
for img_path in image_paths:
    img = Image.open(img_path)
    w, h = img.size
    #  header and footer crop 
    cropped = img.crop((0, int(h*0.05), w, int(h*0.92)))
    text = pytesseract.image_to_string(cropped, lang='ben')
    all_text.append(text)
    print(f"{os.path.basename(img_path)} OCR done")

# Step 3: Word File 
print("\nStep 3: Word file Creating...")
doc = Document()
doc.add_heading('আল-আদাবুল মুফরাদ', 0)
for i, text in enumerate(all_text):
    doc.add_heading(f'page {i+1}', level=2)
    doc.add_paragraph(text)
    doc.add_page_break()
    doc.save(r"C:\Users\Laptop point BD\Desktop\output.docx")
#   doc.save(r"C:\pdf_project\output.docx")
print("output.docx saved")

# Step 4: Excel File 
print("\nStep 4: Excel file Creating...")
wb = Workbook()
ws = wb.active
ws.title = "আদাবুল মুফরাদ"
ws.column_dimensions['A'].width = 15
ws.column_dimensions['B'].width = 100
ws['A1'] = 'page number'
ws['B1'] = 'Text'
for i, text in enumerate(all_text):
    ws.append([f'page {i+1}', text.strip()])
wb.save(r"C:\pdf_project\output.xlsx")
print("output.xlsx saved")

print("\n All work is done!")
print(" C:\\pdf_project\\ Check   output.docx ও output.xlsx files")
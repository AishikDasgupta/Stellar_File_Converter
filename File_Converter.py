import os
import PyPDF2
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import csv
from docx import Document
import pandas as pd
# !!!READ THE README FILE BEFORE USING THE CODE FOR INSTALLING REQUIRED DEPENDENCIES!!!
def convert_pdf_to_csv(pdf_path, csv_path):
    text = ""
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()

    with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Extracted Text'])
        csv_writer.writerow([text])

def convert_image_to_csv(image_path, csv_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img, lang='eng')

    with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Extracted Text'])
        csv_writer.writerow([text])

def convert_to_word(docx_path, text):
    doc = Document()
    doc.add_paragraph(text)

    doc.save(docx_path)

def convert_to_excel(excel_path, text):
    df = pd.DataFrame({'Extracted Text': [text]})
    df.to_excel(excel_path, index=False)

ALLOWED_EXTENSIONS = {'pdf', 'jpg', 'jpeg', 'png'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def main():
    input_file = 'C:/Screenshot from 2023-09-15 14-45-23.png'  # Replace with the input file path (PDF, image, etc.)
    output_file = 'C:/Desktop/python projects/File_Converter/output.doc'  # Specify the desired output file path (CSV, Word, Excel, etc.)

    if input_file.lower().endswith('.pdf'):
        convert_pdf_to_csv(input_file, output_file)
    elif input_file.lower().endswith(('.jpg', '.jpeg', '.png')):
        convert_image_to_csv(input_file, output_file)
    else:
        print("Unsupported file format. Please provide a PDF, JPG, JPEG, or PNG file.")

if __name__ == "__main__":
    main()

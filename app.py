# Step 1: Import necessary libraries
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import os
import google.generativeai as genai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import load_workbook
import streamlit as st

# Streamlit UI to upload PDF files
st.title("Invoice PDF Processing")
st.write("Please Upload the Invoice PDFs")
uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)
pdf_paths = [f.name for f in uploaded_files]

# Ask user for the invoice month
month_options = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
selected_month = st.selectbox("Please enter the invoice month:", month_options)

# Streamlit UI to upload Excel file
st.write("Please Upload the Local Master Excel File")
uploaded_excel = st.file_uploader("Choose an Excel file", type="xlsx")
if uploaded_excel:
    excel_path = uploaded_excel.name
    workbook = load_workbook(uploaded_excel)
    worksheet = workbook.active

# Define functions for processing PDFs and extracting data
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text_data = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        text_data.append(text)
    return text_data

def convert_pdf_to_images_and_ocr(pdf_path):
    images = convert_from_path(pdf_path)
    ocr_results = [pytesseract.image_to_string(image) for image in images]
    return ocr_results

def combine_text_and_ocr_results(text_data, ocr_results):
    combined_results = []
    for text, ocr_text in zip(text_data, ocr_results):
        combined_results.append(text + "\n" + ocr_text)
    combined_text = "\n".join(combined_results)
    return combined_text

def extract_parameters_from_response(response_text):
    def sanitize_value(value):
        return value.strip().replace('"', '').replace(',', '')

    parameters = {
        "PO Number": "NA",
        "Invoice Number": "NA",
        "Invoice Amount": "NA",
        "Invoice Date": "NA",
        "CGST Amount": "NA",
        "SGST Amount": "NA",
        "IGST Amount": "NA",
        "Total Tax Amount": "NA",
        "Taxable Amount": "NA",
        "TCS Amount": "NA",
        "IRN Number": "NA",
        "Receiver GSTIN": "NA",
        "Receiver Name": "NA",
        "Vendor GSTIN": "NA",
        "Vendor Name": "NA",
        "Remarks": "NA",
        "Vendor Code": "NA"
    }
    lines = response_text.splitlines()
    for line in lines:
        for key in parameters.keys():
            if key in line:
                value = sanitize_value(line.split(":")[-1].strip())
                parameters[key] = value
    return parameters

# Set up Google Generative AI client
os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=os.environ["GEMINI_API_KEY"])

prompt = ("the following is OCR extracted text from a single invoice PDF. "
          "Please use the OCR extracted text to give a structured summary. "
          "The structured summary should consider information such as PO Number, Invoice Number, Invoice Amount, Invoice Date, "
          "CGST Amount, SGST Amount, IGST Amount, Total Tax Amount, Taxable Amount, TCS Amount, IRN Number, Receiver GSTIN, "
          "Receiver Name, Vendor GSTIN, Vendor Name, Remarks and Vendor Code. If any of this information is not available or present, "
          "then NA must be denoted next to the value. Please do not give any additional information.")

# Set up Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('secret_key.json', SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open("Health&GlowMasterData")
sheet = spreadsheet.worksheet(selected_month)

# Process each PDF and send data to Google Sheets and Excel
if st.button("Process PDFs"):
    for uploaded_pdf in uploaded_files:
        text_data = extract_text_from_pdf(uploaded_pdf)
        ocr_results = convert_pdf_to_images_and_ocr(uploaded_pdf)
        combined_text = combine_text_and_ocr_results(text_data, ocr_results)

        input_text = f"{prompt}\n\n{combined_text}"

        generation_config = {
            "temperature": 1,
            "top_p": 0.95,
            "top_k": 64,
            "max_output_tokens": 8192,
            "response_mime_type": "text/plain",
        }

        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            generation_config=generation_config,
        )

        chat_session = model.start_chat(history=[])
        response = chat_session.send_message(input_text)

        parameters = extract_parameters_from_response(response.text)

        row_data = [parameters[key] for key in parameters.keys()]
        sheet.append_row(row_data)
        worksheet.append(row_data)

        st.write(f"{uploaded_pdf.name} Structured Summary:")
        for key, value in parameters.items():
            st.write(f"{key}: {value}")

    # Save the updated Excel file
    workbook.save(excel_path)
    st.write("Excel file updated successfully.")
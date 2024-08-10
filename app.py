import os
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
import streamlit as st
from PIL import Image
import google.generativeai as genai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import load_workbook

# Step 1: Define functions for processing PDFs and extracting data
def extract_text_from_pdf(pdf_file):
    with fitz.open(stream=pdf_file.read(), filetype="pdf") as doc:
        text_data = []
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text("text")
            text_data.append(text)
    return text_data

def convert_pdf_to_images_and_ocr(pdf_file):
    images = convert_from_path(pdf_file.name)
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

# Step 2: Set up Google Generative AI client
genai.configure(api_key=st.secrets["gemini_api_key"])

# Define the prompt
prompt = ("the following is OCR extracted text from a single invoice PDF. "
          "Please use the OCR extracted text to give a structured summary. "
          "The structured summary should consider information such as PO Number, Invoice Number, Invoice Amount, Invoice Date, "
          "CGST Amount, SGST Amount, IGST Amount, Total Tax Amount, Taxable Amount, TCS Amount, IRN Number, Receiver GSTIN, "
          "Receiver Name, Vendor GSTIN, Vendor Name, Remarks and Vendor Code. If any of this information is not available or present, "
          "then NA must be denoted next to the value. Please do not give any additional information.")

# Step 3: Set up Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('secret_key.json', SCOPES)
client = gspread.authorize(creds)

# UI for uploading files
st.title("Invoice Processing App")
pdf_files = st.file_uploader("Please upload the Invoice PDFs", type="pdf", accept_multiple_files=True)
selected_month = st.selectbox("Please select the invoice month", [
    "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
])
excel_file = st.file_uploader("Please upload the Local Master Excel File", type="xlsx")

if pdf_files and selected_month and excel_file:
    workbook = load_workbook(excel_file)
    worksheet = workbook.active

    # Google Sheets setup
    spreadsheet = client.open("Health&GlowMasterData")
    sheet = spreadsheet.worksheet(selected_month)

    # Process each PDF and send data to Google Sheets and Excel
    for pdf_file in pdf_files:
        text_data = extract_text_from_pdf(pdf_file)
        ocr_results = convert_pdf_to_images_and_ocr(pdf_file)
        combined_text = combine_text_and_ocr_results(text_data, ocr_results)

        input_text = f"{prompt}\n\n{combined_text}"

        # Create the model configuration
        generation_config = {
            "temperature": 1,
            "top_p": 0.95,
            "top_k": 64,
            "max_output_tokens": 8192,
            "response_mime_type": "text/plain",
        }

        # Initialize the model
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            generation_config=generation_config,
        )

        # Start a chat session
        chat_session = model.start_chat(history=[])

        # Send the combined text as a message
        response = chat_session.send_message(input_text)

        # Extract the relevant data from the response
        parameters = extract_parameters_from_response(response.text)

        # Add data to the Google Sheet
        row_data = [parameters[key] for key in parameters.keys()]
        sheet.append_row(row_data)

        # Add data to the Excel file
        worksheet.append(row_data)

        # Print the structured summary
        st.write(f"\n{pdf_file.name} Structured Summary:\n")
        for key, value in parameters.items():
            st.write(f"{key:20}: {value}")

    # Save the updated Excel file and download it
    workbook.save(excel_file.name)
    st.download_button("Download Updated Excel File", data=open(excel_file.name, "rb").read(), file_name=excel_file.name)






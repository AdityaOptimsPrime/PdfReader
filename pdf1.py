import openai
import pdfplumber
import pandas as pd
import streamlit as st
import openpyxl as xl


# Set your OpenAI API key
openai.api_key = "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"
def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF file using pdfplumber."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    st.write(text)
    return text

def ask_chatgpt_to_parse_data(pdf_text):
    """Uses ChatGPT to parse structured data from PDF text."""
    prompt = f"""
    The following text is extracted from a PDF. Extract structured data into the following fields:
    - Document Number
    - Invoice Date
    - Customer PO
    - Item
    - Units
    - Price

    Text:
    {pdf_text}

    Provide the data in JSON format.
    """
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",  # Replace with the appropriate model
        messages=[
            {"role": "system", "content": "You are an expert in parsing and structuring text data."},
            {"role": "user", "content": prompt}
        ]
    )
    return response['choices'][0]['message']['content']

def save_data_to_excel(json_data, output_path):
    """Converts JSON data to an Excel file."""
    df = pd.read_json(json_data)
    df.to_excel(output_path, index=False)

def deleteData():
        workbook = xl.load_workbook("Data.xlsx")
        sheet = workbook.active  # Select the active sheet
        for row in sheet.iter_rows(min_row=2):  # Skip the header row (row 1)
            for cell in row:
                cell.value = None  # Clear the cell's value
        workbook.save("Data.xlsx")

uploadedFile = st.file_uploader("Upload your PDFs...", type=["pdf"], accept_multiple_files=True)


if uploadedFile is not None:
    st.write("PDF Uploaded Successfully")


motosel = st.button("Click to add data of Motosel vendor")



if motosel:
    deleteData()
    if uploadedFile is not None:
        for uploadedFile1 in uploadedFile:
            with open("temp.pdf","wb") as f:
                f.write(uploadedFile1.read())
            pdf_text = extract_text_from_pdf("temp.pdf")
            structured_data = ask_chatgpt_to_parse_data(pdf_text)
            save_data_to_excel(structured_data, "Data.xlsx")

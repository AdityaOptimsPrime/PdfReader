import pytesseract
from pdf2image import convert_from_bytes
import re
import streamlit as st
from PIL import ImageOps,Image, ImageEnhance, ImageFilter

# Ensure Tesseract is installed and configured properly
# For Windows, specify Tesseract installation path:
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# def preprocess_images(image_path):
#     image = Image.open(image_path).convert("L")  # Convert to grayscale
#     image = image.filter(ImageFilter.SHARPEN)  # Sharpen the image
#     return image

def preprocess_image(image):
    """
    Preprocesses an image for better OCR results.
    - Converts to grayscale.
    - Enhances contrast.
    """
    image = ImageOps.grayscale(image)
    return image

def extract_fields_and_table_from_image(image):
    """
    Extracts structured data (fields and table) from an image using OCR.
    """
    # Preprocess the image for better OCR
    image = preprocess_image(image)
    # image=preprocess_images(image)
    # st.write(image)
    # Perform OCR on the image
    custom_config = r"--psm 6 --oem 3 "  # OCR Engine Mode and Page Segmentation Mode
    text = pytesseract.image_to_string(image, config=custom_config)
    # st.write(text)
    
    # Extract fields using regex
    fields = {}
    fields["Document Number"] = re.search(r"Document Number:\s*(\d+)", text)
    fields["Invoice Date"] = re.search(r"Invoice Date:\s*([\d/]+)", text)
    fields["Customer PO"] = re.search(r"Customer PO:\s*(.+)", text)
    
    # Extract table data (Items, Units, and Price)
    table_data = []
    lines = text.splitlines()
    st.write(lines)
    table_started = False
    for line in lines:
        # if "Description" in line:  # Start of table
        #     table_started = True
        #     continue
        if table_started and line.strip():  # Capture table rows
            row = line.split()
            st.write(row)
            if len(row) >= 5:
                table_data.append({
                    "Item": row[1],
                    "Units": row[0],
                    "Price": row[-2]
                })

    # Clean and finalize fields
    fields = {key: match.group(1).strip() if match else None for key, match in fields.items()}
    fields["Items"] = table_data

    return fields

def extract_from_pdf(pdf_bytes):
    """
    Extracts structured data from an image-based PDF.
    """
    extracted_data = []
    images = convert_from_bytes(pdf_bytes, dpi=300)  # Convert PDF pages to images

    for i, image in enumerate(images):
        if i == 0:  # Skip the first page
            continue
        page_data = extract_fields_and_table_from_image(image)
        extracted_data.extend(page_data)
    
    return extracted_data

# Streamlit App
st.title("Image-Based PDF Table Extractor")
st.write("Upload an image-based PDF, and the app will extract structured data from it.")

uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")

if uploaded_file is not None:
    try:
        st.write("Processing PDF...")
        pdf_bytes = uploaded_file.read()
        extracted_data = extract_from_pdf(pdf_bytes)
        
        if extracted_data:
            st.success("Data extracted successfully!")
            st.json(extracted_data)
        else:
            st.warning("No data extracted. Ensure the PDF contains recognizable text.")
    except Exception as e:
        st.error(f"An error occurred: {e}")

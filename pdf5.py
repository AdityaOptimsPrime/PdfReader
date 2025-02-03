from PIL import Image
import pytesseract
import streamlit as st
from pdf2image import convert_from_bytes
import io
import re

# Specify the path to the Tesseract executable (for Windows)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Streamlit app title
st.title("Extract Specific Fields from PDF")

# File uploader for PDFs
uploaded_file = st.file_uploader("Upload a PDF file...", type=["pdf"])

# Function to extract specific fields using regex
def extract_fields(text):
    fields = {
        "Document Number": None,
        "Invoice Date": None,
        "Units": None,
        "Item": None,
        "Price": None,
    }

    # Updated regex patterns for better matching
    patterns = {
        "Document Number": r"(?i)(?:document|doc|invoice)[.\s]*(?:number|no|#)?[:\s]*(\w+)",
        "Invoice Date": r"(?i)(?:invoice|date)[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})",
        "Units": r"(?i)(?:units?|quantity)[:\s]*(\d+)",
        "Item": r"(?i)(?:item|description|product)[:\s]*([A-Za-z0-9\s\-]+)",
        "Price": r"(?i)(?:price|amount|total)[:\s]*[$]?\s*(\d+(?:\.\d{2})?)",
    }

    # Extract fields using regex
    for field, pattern in patterns.items():
        matches = re.finditer(pattern, text, re.MULTILINE)
        for match in matches:
            if match and not fields[field]:  # Take the first match if field is empty
                fields[field] = match.group(1).strip()

    return fields

# Check if a PDF is uploaded
if uploaded_file is not None:
    st.write("PDF Uploaded Successfully")

    # Button to trigger text extraction
    if st.button("Extract Fields from PDF"):
        try:
            # Convert PDF to images
            images = convert_from_bytes(uploaded_file.read())
            
            if not images:
                st.warning("No pages found in the PDF.")
            else:
                st.write(f"Processing {len(images)} page(s) from the PDF.")

                # Extract text from each image and process fields
                all_fields = {}
                for i, image in enumerate(images):
                    # Enhance image quality for better OCR
                    enhanced_image = image.convert('L')  # Convert to grayscale
                    
                    # Display the processed image
                    st.image(enhanced_image, caption=f"Page {i+1}", use_column_width=True)
                    
                    # Extract text with improved OCR configuration
                    extracted_text = pytesseract.image_to_string(
                        enhanced_image,
                        config='--psm 6 --oem 3'  # Page segmentation mode 6 (Assume uniform block of text)
                    )
                    
                    # Display raw extracted text for debugging
                    with st.expander(f"Raw Extracted Text (Page {i+1})"):
                        st.text(extracted_text)
                    
                    # Extract fields from the text
                    fields = extract_fields(extracted_text)

                    # Combine fields from all pages
                    for key, value in fields.items():
                        if value and key not in all_fields:  # Only add if value exists and field not already found
                            all_fields[key] = value

                # Display the extracted fields
                st.subheader("Extracted Fields:")
                if all_fields:
                    for field, value in all_fields.items():
                        st.write(f"{field}: {value}")
                else:
                    st.warning("No fields could be extracted. Please check if the PDF contains the required information.")

        except Exception as e:
            st.error(f"An error occurred: {e}")
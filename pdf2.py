import fitz  # PyMuPDF
import requests
import streamlit as st

# Function to extract text from PDF
def extract_pdf_data(pdf_path):
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)
    
    # Initialize an empty string to store text
    extracted_text = ""
    
    # Iterate through each page in the PDF
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)  # Load the page
        extracted_text += page.get_text("text")  # Extract text from the page
        print(page)
    return extracted_text

# Function to send extracted text to Gemini API (or any language model API)
def query_gemini(prompt, text):
    url = "https://gemini.googleapis.com/v1/completions"  # Example URL for Gemini
    headers = {
        "Authorization": "Bearer AIzaSyAO3BoIA1sckavAW97KEKcf_3SEFUddHHM",  # Replace with your actual API key
        "Content-Type": "application/json"
    }
    
    # Craft a prompt for Gemini to extract specific fields from the text
    full_prompt = f"""
    You are an AI assistant. I will provide you with an invoice text, and I need you to extract specific fields from it. The fields to extract are:
    
    1. Document Number
    2. Invoice Date
    3. Customer PO
    4. Item
    5. Units
    6. Price

    The invoice text is as follows:

    {text}

    Please extract the values for these fields and return them in the following format:
    
    Document Number: [value]
    Invoice Date: [value]
    Customer PO: [value]
    Item: [value]
    Units: [value]
    Price: [value]
    """
    
    # Create the payload for the API request
    payload = {
        "model": "gemini-1.5-flash",  # Replace with the actual model name for Gemini
        "prompt": full_prompt,
        "max_tokens": 500,
        "temperature": 0.7
    }
    
    # Make the POST request to Gemini API
    response = requests.post(url, json=payload, headers=headers)
    
    # Return the response if successful, or error message
    if response.status_code == 200:
        return response.json()["choices"][0]["text"]
    else:
        return f"Error: {response.status_code}"

# Streamlit app interface
def main():
    st.title("PDF Text Extraction and Gemini API Integration")
    
    # Upload file
    uploaded_file = st.file_uploader("Upload a PDF", type="pdf")
    
    if uploaded_file is not None:
        # Extract the text from the PDF
        pdf_text = extract_pdf_data("temp.pdf")
        
        # Display a portion of the extracted text
        st.subheader("Extracted Text Preview:")
        st.text(pdf_text[:1000])  # Show the first 1000 characters of the text
        
        # Send extracted text to Gemini to extract specific fields
        st.subheader("Extracted Fields from Gemini:")
        response = query_gemini("Extract Invoice Information", pdf_text)
        
        # Display the response from Gemini
        st.text(response)
        
        st.success("PDF Processing Completed!")

if __name__ == "__main__":
    main()

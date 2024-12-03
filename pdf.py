import streamlit as st
import PyPDF2 as pPyPDFReaderdf
import openpyxl as xl
import re
import pdfplumber as pdfp
import pandas as pd
import camelot


url="https://docs.google.com/spreadsheets/d/1xZvNRfY2fq6ixt8qDhymX1drmtQkJAXV1-7kHlpydJY/edit?usp=sharing"


def aptPdf(filePath):
    with pdfp.open(filePath) as file:
        table1=[]
        for i in file.pages:
            extracted_tables = i.extract_tables()
            tables=extracted_tables[0][1][1]
            data1=tables.split(":")[1]
            break

        for i in file.pages:
            extracted_tables=i.extract_tables()
            data2=(extracted_tables[1][4][0])
            data3=(extracted_tables[1][4][1])
            break

        for i in file.pages:
            extracted_tables=i.extract_tables()
            table1.append(extracted_tables[2])
            break
        
        
        itemName=[]
        itemQuantity=[]
        for row in table1:
            for cell in row:
                itemName.append(cell[0])
                itemQuantity.append(cell[2])
        itemName=itemName[1:]
        itemQuantity=itemQuantity[1:]
        itemMap = dict(zip(itemName, itemQuantity))

        return data1,data2,data3,itemMap
    

def bandoPdf(filePath):

    with pdfp.open(filePath) as file:
        tables = []
        allDicts = []
        for i in file.pages:
            extracted_tables = i.extract_tables()
            tables.extend(extracted_tables)

        if tables:
            last_table = tables[-1]  # Get the last table
            df = pd.DataFrame(last_table)  # Convert it into a DataFrame
            df1=df.iloc[1,2]
            df = df.iloc[:, [1, 8]]
            df = df.dropna()
            
            codeList = []
            product_codes = df.iloc[1:, 0].tolist()  # Get the first column as a list
            quantities = df.iloc[1:, 1].tolist()  # Get the second column as a list

            # Iterate over product_codes and split each product code
            for code in product_codes:
                print(type(code))
                split_code = code.split(" ")
                codeList.append(split_code[0])  # Assuming you only want the first part (e.g., '6PK1145')

            # Ensure that all elements in codeList are strings, not lists
            if any(isinstance(item, list) for item in codeList):
                st.write("Error: codeList contains lists instead of strings")
                return
            product_quantity_dict = dict(zip(codeList, quantities))
            allDicts.append(product_quantity_dict)
        else:
            st.write("No tables found in the PDF")
    return allDicts

def bandoPdf1(filePath1):
    tables = camelot.read_pdf(filepath=filePath1, flavor='stream', pages='all')
    allDicts = []
    for table in tables:
        df = table.df
        # Identify the correct columns based on their headers
        # Replace 'Product and Description' and 'Ship Qty' with their exact column names
        # Remove blank rows if necessary and reset the index
        if df.shape[1]>=4:
            df = df[[df.columns[1], df.columns[4]]]  # Adjust indices for correct columns
            df.columns = ['Product and Description', 'Ship Qty']  # Rename columns for clarity
            df = df[
                (df['Product and Description'] != 'Product and Description') &
                (df['Ship Qty'] != 'Ship Qty') &
                (df['Product and Description'] != '') &
                (df['Ship Qty'] != '')
            ]
            st.write(df)
            # Display the filtered data (or further process it)
            data_dict = dict(zip(df['Product and Description'], df['Ship Qty']))
            allDicts.append(data_dict)
    return allDicts

def comonBandoPdf(filePath1,pageNumber):
    tables = camelot.read_pdf(filepath=filePath1)
    df = tables[3].df
    df = df[[df.columns[1]]]
    PoNumber= df.iloc[1, 0]
    with open(filePath1, 'rb') as file:
        pdfReader = pPyPDFReaderdf.PdfReader(file)  # Properly count the number of pages
        text = ""
        for page_num in range(pageNumber):
            page = pdfReader.pages[page_num]
            text += page.extract_text()
    # Patterns for extracting invoice data
    invoiceDatePattern = r"Due Date\s*(\d{1,2}/\d{1,2}/\d{2})"
    invoiceNoPattern = r"Invoice\s*(\d+-\d+)"
    
    # Search for patterns in the extracted text
    invoiceDate = re.search(invoiceDatePattern, text)
    invoiceNo = re.search(invoiceNoPattern, text)
    
    if invoiceDate and invoiceNo:
        invoiceDate = invoiceDate.group(1)
        invoiceNo = invoiceNo.group(1)
        return PoNumber, invoiceDate, invoiceNo
    else:
        return None, None, None

def bestBuyPdf(filePath):
    with pdfp.open(filePath) as pdf1:
        shipPartList = []
        
        for page in pdf1.pages:
            text = page.extract_text()
            lines = text.split('\n')
            
            for line in lines:
                match = re.match(r"^\s*(\d+)\s+\d+\s+[A-Z]+\s+([A-Z0-9]+)", line)
                if match:
                    ship_qty = match.group(1)
                    part_number = match.group(2)
                    shipPartList.append((ship_qty, part_number))
        
    with open(filePath,'rb') as file:
        pdfReader=pPyPDFReaderdf(file)
        pageNumber= len(pdfReader.pages)
        text=""
        for page_num in range(pageNumber):
            page=pdfReader.pages[page_num]
            text+=page.extract_text()

    poNoPattern = r"P.O.No:\s*(\d{9})"
    invoiceDatePattern = r" Invoice Date:\s*((?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)/\d{1,2}/\d{4})"
    invoiceNoPattern = r"Invoice No:\s*(\d+)"
   
    poNo = re.search(poNoPattern, text)
    invoiceDate = re.search(invoiceDatePattern, text)
    invoiceNo = re.search(invoiceNoPattern, text)
    if poNo and invoiceDate and invoiceNo:
        poNo = poNo.group(1)
        invoiceDate = invoiceDate.group(1)
        invoiceNo=invoiceNo.group(1)  
        return poNo, invoiceDate,invoiceNo,shipPartList
    else:
        return None, None,None


def densoPdf(filePath):
    item_shipped_dict = {}
    with pdfp.open(filePath) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            # Extract raw text from the page
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    # Use regex to extract SHIPPED and ITEM NO. data
                    match = re.match(r"^\s*(\d+)\s+\d+\s+([\d-]+)", line)
                    if match:
                        shipped = match.group(1)
                        item_no = match.group(2)
                        item_shipped_dict[item_no] = int(shipped)
            else:
                st.write(f"Page {page_number} contains no extractable text.")
    # st.write(item_shipped_dict)

    with pdfp.open(filePath) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                # Filter required columns
                invoiceDate=df.iloc[:,1].tolist()
                customerPoNo=df.iloc[:,4].tolist()
            break
    
    with open(filePath,'rb') as file:
        pdfReader=pPyPDFReaderdf.PdfReader(file)
        pageNumber= len(pdfReader.pages)
        text=""
        for page_num in range(pageNumber):
            page=pdfReader.pages[page_num]
            text+=page.extract_text()
            # st.write(text)
    invoicePattern=r"Invoice\s*(\d+)"
    invoice=re.search(invoicePattern,text)
    if invoice:
        invoice=invoice.group(1)
    st.write(invoice)
    
    return invoice,invoiceDate[0],customerPoNo[0],item_shipped_dict

def extractText(filePath):
    with open(filePath,'rb') as file:
        pdfReader=pPyPDFReaderdf.PdfReader(file)
        pageNumber= len(pdfReader.pages)
        text=""
        for page_num in range(pageNumber):
            page=pdfReader.pages[page_num]
            text+=page.extract_text()
            st.write(text)
    # poNoPattern = r"P.O.No:\s*(\d{9})"
    # invoiceDatePattern = r"Due Date\s*(\d{1,2}/\d{1,2}/\d{2})"
    # invoiceNoPattern = r"Invoice\s*(\d+-\d+)"
   
    # # poNo = re.search(poNoPattern, text)
    # invoiceDate = re.search(invoiceDatePattern, text)
    # invoiceNo = re.search(invoiceNoPattern, text)
    # # st.write(poNo)
    # st.write(invoiceDate)
    # st.write(invoiceNo)
    # if invoiceDate and invoiceNo:
    #     # poNo = poNo.group(1)
    #     invoiceDate = invoiceDate.group(1)
    #     invoiceNo=invoiceNo.group(1)
       
    #     return invoiceDate,invoiceNo
    # else:
    #     return  None,None
    

def display_pdf_content(filePath):

     with open(filePath, 'rb') as file:
        pdf_reader = pPyPDFReaderdf(file)
        num_pages = len(pdf_reader.pages)
        text = []
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()
            text.append(page_text)
        return text



def addToExcel (*args):
    workbook = xl.load_workbook("Data.xlsx")
    sheet = workbook.active
    sheet.append(args)
    workbook.save("Data.xlsx")
    # conn = st.connection("gsheets", type=GSheetsConnection)
    # data = conn.read(spreadsheet=url)
    # new_row = pd.DataFrame([args], columns=data.columns)
    
    # Concatenate the new row with the existing data
    # updated_data = pd.concat([data, new_row], ignore_index=True)
    
    # Write the updated data back to the Google Sheet
    # conn.update(data=updated_data)
    # st.dataframe(data)

def get_page_count(pdf_file):
    # Load the PDF file
    reader = pPyPDFReaderdf.PdfReader(pdf_file)
    # Get the total number of pages
    page_count = len(reader.pages)
    return page_count


uploadedFile = st.file_uploader("Upload your PDFs...", type=["pdf"], accept_multiple_files=True)


if uploadedFile is not None:
    st.write("PDF Uploaded Successfully")

submit1 = st.button("Click to add data in spreed sheet")
apt = st.button("Click to add data of APT vendor")
bando = st.button("Click to add data of bando vendor")
bestBuy = st.button("Click to add data of Best Buy Distributor-new vendor")
denso = st.button("Click to add data of DENSO vendor")



if bando:
    if uploadedFile is not None:
        for uploadedFile1 in uploadedFile:
            with open("temp.pdf","wb") as f:
                f.write(uploadedFile1.read())
            count=get_page_count("temp.pdf")
            PoNumber,invoiceDate,invoiceNo=comonBandoPdf("temp.pdf",count)
            if count<2:
                product_quantity_dict=bandoPdf("temp.pdf")
            else:
                product_quantity_dict=bandoPdf1("temp.pdf")
            for i, product_quantity_dict in enumerate(product_quantity_dict, 1):
                for key, value in product_quantity_dict.items():
                    addToExcel(PoNumber,invoiceDate,invoiceNo,key,value)

if apt:
    if uploadedFile is not None:
        with open("temp.pdf","wb") as f:
            f.write(uploadedFile.read())
        data1,data2,data3,itemMap=aptPdf("temp.pdf")
        for key, value in itemMap.items():
            addToExcel(data1,data2,data3,key, value)



if bestBuy:
    if uploadedFile is not None:
        with open("temp.pdf","wb") as f:
            f.write(uploadedFile.read())
        poNo, invoiceDate,invoiceNo,shipPartList=bestBuyPdf("temp.pdf")
        for ship_qty, part_number in shipPartList:
            addToExcel(poNo, invoiceDate,invoiceNo, ship_qty, part_number)


if denso:
    if uploadedFile is not None:
        for uploadedFile1 in uploadedFile:
            with open("temp.pdf","wb") as f:
                f.write(uploadedFile1.read())
            data1,data2,data3,itemMap=densoPdf("temp.pdf")
            for key, value in itemMap.items():
                addToExcel(data1,data2,data3,key, value)
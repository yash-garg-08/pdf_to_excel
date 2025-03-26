import re
import fitz
from openpyxl import Workbook

def process_pdf_and_extract_data(input_pdf, output_file):
    
    #opening the input pdf file
    pdf_document = fitz.open(input_pdf)

    #creating a workbook to create and access excel file
    excel_workbook = Workbook()
    active_sheet = excel_workbook.active
    
    extracted_content = []
    
    for page in pdf_document:
        #getting a page at a time from the pdf
        page_text = page.get_text("text")

        #splitting the text based on each lone break
        lines = page_text.split("\n")
        
        for line in lines:
            #splitting the line based on the multiple spaces with tabs
            while "  " in line:
              line = line.replace("  ", "\t")
            structured_data = line.split("\t")
            
            if structured_data:
                extracted_content.append(structured_data)
    
    #entering each value of row in excel one by one
    for entry in extracted_content:
        active_sheet.append(entry)
    
    #saving the excel sheet
    excel_workbook.save(output_file)

pdf_source = "/content/pdf_reader/test3 (1).pdf"  
excel_destination = "/content/pdf_reader/output.xlsx"
process_pdf_and_extract_data(pdf_source, excel_destination)

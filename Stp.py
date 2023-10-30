import streamlit as st
import os
import pythoncom
import win32com.client as win32
import win32com.client
import gdown


url = "https://docs.google.com/spreadsheets/d/1nXbSZm2VwElRHfnTilPmURnsXTx9q6p1/edit?usp=sharing&ouid=113403734626643599045&rtpof=true&sd=true"
output = 'C:\\Sunstar\\sks.xlsx'
os.makedirs(os.path.dirname(output), exist_ok=True)
if not os.path.exists(output):
    # File doesn't exist, so download it
    gdown.download(url, output, quiet=False)



# Set up Streamlit app title and layout
st.title("File Fill and Save as PDF")

# Create input fields
date = st.text_input("Date")
place = st.text_input("Place")
branch = st.text_input("Branch")
amount = st.text_input("Loan Amount")
name = st.text_input("Name")
roi = st.text_input("ROI")

sunstar_folder_path = 'C:\\Download_Agreement'

if not os.path.exists(sunstar_folder_path):
            os.makedirs(sunstar_folder_path)

# Check if the user provided all required input
if date and place and branch and amount and name and roi:
    # Download the Excel file if it doesn't exist
    file_path=(sunstar_folder_path)
    def save_data():
        try:
            pythoncom.CoInitialize()
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(output)
            ws = workbook.Worksheets('DP note')

            ws.Unprotect('sunstar')
            ws.Cells(3, 12).Value = date
            ws.Cells(4, 12).Value = place
            ws.Cells(10, 12).Value = branch
            ws.Cells(12, 12).Value = amount
            ws.Cells(16, 12).Value = name
            ws.Cells(16, 15).Value = roi

            columns_to_hide = ['L', 'M', 'N', 'O']

            for column in columns_to_hide:
                ws.Columns(column).Hidden = True

            ws.Protect('sunstar')
            workbook.Save()
            excel.Application.Quit()
            st.write("Data has been saved to the Excel file.")
        except Exception as e:
            st.write(f"An error occurred: {str(e)}")

        
    # Define a function to convert Excel to PDF
    def excel_to_pdf():
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(output)
            pdf_file_path = os.path.join(sunstar_folder_path, 'agreement.pdf')
            workbook.ExportAsFixedFormat(0, pdf_file_path, Quality=0)
            workbook.Close(SaveChanges=False)
            excel.Quit()
            st.write("PDF has been generated and is ready for download.")
        except Exception as e:
            st.write(f"An error occurred: {str(e)}")

    # Create buttons for saving to Excel and converting to PDF
    if st.button("SAVE"):
        save_data()

    if st.button("Download PDF"):
        excel_to_pdf()

    result_label = st.empty()
else:
    st.warning("Please fill in all required fields.")

import xlwt
import pdfplumber
import re
import zipfile
import os

def extract_cv_info(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text()

        # Extract email using regex
        email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        emails = re.findall(email_regex, text)
        email = emails[0] if emails else "No email found"

        # Extract phone number using regex
        phone_regex = r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]'
        phones = re.findall(phone_regex, text)
        phone = phones[0] if phones else "No phone number found"

        # Remove email and phone from the text
        text = re.sub(email_regex, "", text)
        text = re.sub(phone_regex, "", text)

        return {
            'email': email,
            'phone': phone,
            'text': text.strip(),
        }
    except pdfplumber.PDFSyntaxError as e:
        print(f"PDF Syntax Error: {str(e)}")
        # Return a default value or handle the error as needed
        return {
            'email': "",
            'phone': "",
            'text': "",
        }
    except Exception as e:
        print(f"An error occurred while processing the PDF: {str(e)}")
        # Return a default value or handle the error as needed
        return {
            'email': "",
            'phone': "",
            'text': "",
        }

def create_excel(cv_data):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('CV Information')
    headers = ['Email', 'Phone', 'Text']
    for col, header in enumerate(headers):
        sheet.write(0, col, header)

    for row, data in enumerate(cv_data, start=1):
        sheet.write(row, 0, data['email'])
        sheet.write(row, 1, data['phone'])
        sheet.write(row, 2, data['text'])

    excel_file = 'cv_info.xls'
    workbook.save(excel_file)
    return excel_file

def extract_zip_files(zip_file):
    extracted_files = []
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall('temp_zip')
        for filename in zip_ref.namelist():
            extracted_files.append(os.path.join('temp_zip', filename))
    return extracted_files

def cleanup_temp_files():
    if os.path.exists('temp_zip'):
        for root, dirs, files in os.walk('temp_zip', topdown=False):
            for file in files:
                try:
                    file_path = os.path.join(root, file)
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"Error cleaning up temp file: {str(e)}")
            for dir in dirs:
                try:
                    dir_path = os.path.join(root, dir)
                    if os.path.isdir(dir_path):
                        os.rmdir(dir_path)
                except Exception as e:
                    print(f"Error cleaning up temp directory: {str(e)}")
        os.rmdir('temp_zip')

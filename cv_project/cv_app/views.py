from django.shortcuts import render
from django.http import HttpResponse
from .utils import extract_cv_info, create_excel, extract_zip_files, cleanup_temp_files

def upload_cv(request):
    if request.method == 'POST':
        zip_file = request.FILES['zip_file']
        if zip_file.name.endswith('.zip'):
            # Extract files from the zip
            extracted_files = extract_zip_files(zip_file)
            
            # Process each extracted PDF
            cv_data = []
            for cv_file in extracted_files:
                cv_info = extract_cv_info(cv_file)
                cv_data.append(cv_info)

            # Create Excel file
            excel_file = create_excel(cv_data)
            
            # Clean up temporary files
            cleanup_temp_files()

            # Serve the Excel file as a response
            response = HttpResponse(content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename="cv_info.xls"'
            excel_file_content = open(excel_file, 'rb').read()
            response.write(excel_file_content)
            return response
        else:
            return HttpResponse("Please upload a .zip file.")
    else:
        return render(request, 'upload.html')

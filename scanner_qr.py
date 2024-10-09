# import cv2
# from pyzbar.pyzbar import decode
# import os
# import requests
# import pandas as pd

# def scan_qr_code(image_path):
#     img = cv2.imread(image_path)
#     qr_codes = decode(img)
    
#     if qr_codes:
#         for qr_code in qr_codes:
#             return qr_code.data.decode('utf-8')
#     return None

# def scan_certificates_in_folder(folder_path):
#     scanned_results = []
    
#     for certificate_file in os.listdir(folder_path):
#         certificate_path = os.path.join(folder_path, certificate_file)
        
#         if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
#             qr_data = scan_qr_code(certificate_path)
            
#             if qr_data:
#                 scanned_results.append((certificate_file, qr_data))
#             else:
#                 scanned_results.append((certificate_file, "No QR code found"))
    
#     return scanned_results

# def is_url_valid(url):
#     try:
#         response = requests.head(url, allow_redirects=True, timeout=5)
#         return response.status_code == 200
#     except requests.RequestException:
#         return False

# def save_results_to_excel(results, excel_path):
#     df = pd.DataFrame(results, columns=['Certificate', 'QR Data', 'URL Status'])
#     df.to_excel(excel_path, index=False)

# # Folder containing certificates
# certificates_folder = 'certificates/'

# # Scan all certificates in the folder
# scanned_data = scan_certificates_in_folder(certificates_folder)

# # Prepare the data to be saved in Excel
# results = []
# for certificate, qr_data in scanned_data:
#     if qr_data != "No QR code found":
#         valid = is_url_valid(qr_data)
#         validity_status = "Valid" if valid else "Invalid"
#         results.append((certificate, qr_data, validity_status))
#     else:
#         results.append((certificate, qr_data, "N/A"))

# # Save the results to an Excel file
# excel_file_path = 'scanned_results.xlsx'
# save_results_to_excel(results, excel_file_path)

# print(f"Results saved to {excel_file_path}")

































# import cv2
# from pyzbar.pyzbar import decode
# import os
# import requests
# import pandas as pd
# import pytesseract
# from pdf2image import convert_from_path

# # Function to scan QR code from an image
# def scan_qr_code(image_path):
#     img = cv2.imread(image_path)
#     qr_codes = decode(img)
    
#     if qr_codes:
#         for qr_code in qr_codes:
#             return qr_code.data.decode('utf-8')
#     return None

# # Function to extract name and registration number using pytesseract
# def extract_name_and_registration(image_path):
#     img = cv2.imread(image_path)
#     text = pytesseract.image_to_string(img)
    
#     # Find the name and registration number in the extracted text
#     name = None
#     registration_number = None
    
#     for line in text.splitlines():
#         if "Name" in line:
#             name = line.split(":")[-1].strip()
#         if "Registration" in line:
#             registration_number = line.split(":")[-1].strip()
    
#     return name, registration_number

# # Function to scan a certificate and handle both image and PDF formats
# def process_certificate(file_path, file_name):
#     if file_name.endswith(('.png', '.jpg', '.jpeg')):
#         qr_data = scan_qr_code(file_path)
#         name, registration_number = extract_name_and_registration(file_path)
#     elif file_name.endswith('.pdf'):
#         # Convert PDF to images
#         images = convert_from_path(file_path)
#         qr_data = None
#         name, registration_number = None, None
#         for img in images:
#             img_path = "temp_image.png"  # Temporary image for processing
#             img.save(img_path, 'PNG')
            
#             # Try scanning QR code from the image
#             qr_data = scan_qr_code(img_path)
#             if qr_data:
#                 name, registration_number = extract_name_and_registration(img_path)
#                 os.remove(img_path)  # Clean up temp image
#                 break  # Stop after first page if QR code is found
#     else:
#         return None, None, None

#     return qr_data, name, registration_number

# # Function to validate URL
# def is_url_valid(url):
#     try:
#         response = requests.head(url, allow_redirects=True, timeout=10)
#         return response.status_code == 200
#     except requests.RequestException:
#         return False

# # Function to scan all certificates in the folder and collect data
# def scan_certificates_in_folder(folder_path):
#     scanned_results = []
    
#     for certificate_file in os.listdir(folder_path):
#         certificate_path = os.path.join(folder_path, certificate_file)
        
#         if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
#             qr_data, name, registration_number = process_certificate(certificate_path, certificate_file)
            
#             if qr_data:
#                 scanned_results.append((certificate_file, qr_data, name or "Name not found", registration_number or "Registration number not found"))
#             else:
#                 scanned_results.append((certificate_file, "No QR code found", name or "Name not found", registration_number or "Registration number not found"))
    
#     return scanned_results

# # Function to save results to Excel
# def save_results_to_excel(results, excel_path):
#     df = pd.DataFrame(results, columns=['Certificate', 'QR Data', 'Name', 'Registration Number', 'URL Status'])
#     df.to_excel(excel_path, index=False)

# # Folder containing certificates
# certificates_folder = 'certificates/'

# # Scan all certificates in the folder
# scanned_data = scan_certificates_in_folder(certificates_folder)

# # Prepare the data to be saved in Excel
# results = []
# for certificate, qr_data, name, registration_number in scanned_data:
#     if qr_data != "No QR code found":
#         valid = is_url_valid(qr_data)
#         validity_status = "Valid" if valid else "Invalid"
#         results.append((certificate, qr_data, name, registration_number, validity_status))
#     else:
#         results.append((certificate, qr_data, name, registration_number, "N/A"))

# # Save the results to an Excel file
# excel_file_path = 'scanned_results.xlsx'
# save_results_to_excel(results, excel_file_path)

# print(f"Results saved to {excel_file_path}")




































# import cv2
# from pyzbar.pyzbar import decode
# import os
# import requests
# import pandas as pd
# import pytesseract
# from pdf2image import convert_from_path

# # Ensure Tesseract is in the PATH or specify the exact location
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update this path if needed

# # Function to scan QR code from an image
# def scan_qr_code(image_path):
#     img = cv2.imread(image_path)
#     qr_codes = decode(img)
    
#     if qr_codes:
#         for qr_code in qr_codes:
#             return qr_code.data.decode('utf-8')
#     return None

# # Function to extract name and registration number using pytesseract
# def extract_name_and_registration(image_path):
#     img = cv2.imread(image_path)
#     text = pytesseract.image_to_string(img)
    
#     # Find the name and registration number in the extracted text
#     name = None
#     registration_number = None
    
#     for line in text.splitlines():
#         if "Name" in line:
#             name = line.split(":")[-1].strip()
#         if "Registration" in line:
#             registration_number = line.split(":")[-1].strip()
    
#     return name, registration_number

# # Function to scan a certificate and handle both image and PDF formats
# def process_certificate(file_path, file_name):
#     if file_name.endswith(('.png', '.jpg', '.jpeg')):
#         qr_data = scan_qr_code(file_path)
#         name, registration_number = extract_name_and_registration(file_path)
#     elif file_name.endswith('.pdf'):
#         # Convert PDF to images
#         images = convert_from_path(file_path)
#         qr_data = None
#         name, registration_number = None, None
#         for img in images:
#             img_path = "temp_image.png"  # Temporary image for processing
#             img.save(img_path, 'PNG')
            
#             # Try scanning QR code from the image
#             qr_data = scan_qr_code(img_path)
#             if qr_data:
#                 name, registration_number = extract_name_and_registration(img_path)
#                 os.remove(img_path)  # Clean up temp image
#                 break  # Stop after first page if QR code is found
#     else:
#         return None, None, None

#     return qr_data, name, registration_number

# # Function to validate URL
# def is_url_valid(url):
#     try:
#         response = requests.head(url, allow_redirects=True, timeout=10)
#         return response.status_code == 200
#     except requests.RequestException:
#         return False

# # Function to scan all certificates in the folder and collect data
# def scan_certificates_in_folder(folder_path):
#     scanned_results = []
    
#     for certificate_file in os.listdir(folder_path):
#         certificate_path = os.path.join(folder_path, certificate_file)
        
#         if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
#             qr_data, name, registration_number = process_certificate(certificate_path, certificate_file)
            
#             if qr_data:
#                 scanned_results.append((certificate_file, qr_data, name or "Name not found", registration_number or "Registration number not found"))
#             else:
#                 scanned_results.append((certificate_file, "No QR code found", name or "Name not found", registration_number or "Registration number not found"))
    
#     return scanned_results

# # Function to save results to Excel
# def save_results_to_excel(results, excel_path):
#     df = pd.DataFrame(results, columns=['Certificate', 'QR Data', 'Name', 'Registration Number', 'URL Status'])
#     df.to_excel(excel_path, index=False)

# # Folder containing certificates
# certificates_folder = 'certificates/'

# # Scan all certificates in the folder
# scanned_data = scan_certificates_in_folder(certificates_folder)

# # Prepare the data to be saved in Excel
# results = []
# for certificate, qr_data, name, registration_number in scanned_data:
#     if qr_data != "No QR code found":
#         valid = is_url_valid(qr_data)
#         validity_status = "Valid" if valid else "Invalid"
#         results.append((certificate, qr_data, name, registration_number, validity_status))
#     else:
#         results.append((certificate, qr_data, name, registration_number, "N/A"))

# # Save the results to an Excel file
# excel_file_path = 'scanned_results.xlsx'
# save_results_to_excel(results, excel_file_path)

# print(f"Results saved to {excel_file_path}")


































import cv2
from pyzbar.pyzbar import decode
import os
import requests
import pandas as pd
import pytesseract
from pdf2image import convert_from_path


# # Ensure Tesseract is in the PATH or specify the exact location
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to scan QR code from an image
def scan_qr_code(image_path):
    img = cv2.imread(image_path)
    qr_codes = decode(img)
    
    if qr_codes:
        for qr_code in qr_codes:
            return qr_code.data.decode('utf-8')
    return None

# Function to extract name and registration number using pytesseract
def extract_name_and_registration(image_path):
    img = cv2.imread(image_path)
    text = pytesseract.image_to_string(img)
    
    # Find the name and registration number in the extracted text
    name = None
    registration_number = None
    
    for line in text.splitlines():
        if "Name" in line:
            name = line.split(":")[-1].strip()
        if "Registration" in line:
            registration_number = line.split(":")[-1].strip()
    
    return name, registration_number

# Function to scan a certificate and handle both image and PDF formats
def process_certificate(file_path, file_name):
    if file_name.endswith(('.png', '.jpg', '.jpeg')):
        qr_data = scan_qr_code(file_path)
        name, registration_number = extract_name_and_registration(file_path)
    elif file_name.endswith('.pdf'):
        # Convert PDF to images
        images = convert_from_path(file_path)
        qr_data = None
        name, registration_number = None, None
        for img in images:
            img_path = "temp_image.png"  # Temporary image for processing
            img.save(img_path, 'PNG')
            
            # Try scanning QR code from the image
            qr_data = scan_qr_code(img_path)
            if qr_data:
                name, registration_number = extract_name_and_registration(img_path)
                os.remove(img_path)  # Clean up temp image
                break  # Stop after first page if QR code is found
    else:
        return None, None, None

    return qr_data, name, registration_number

# Function to validate URL
def is_url_valid(url):
    try:
        response = requests.head(url, allow_redirects=True, timeout=10)
        return response.status_code == 200
    except requests.RequestException:
        return False

# Function to scan all certificates in the folder and collect data
def scan_certificates_in_folder(folder_path):
    scanned_results = []
    
    for certificate_file in os.listdir(folder_path):
        certificate_path = os.path.join(folder_path, certificate_file)
        
        if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
            qr_data, name, registration_number = process_certificate(certificate_path, certificate_file)
            
            if qr_data:
                scanned_results.append((certificate_file, qr_data, name or "Name not found", registration_number or "Registration number not found"))
            else:
                scanned_results.append((certificate_file, "No QR code found", name or "Name not found", registration_number or "Registration number not found"))
    
    return scanned_results

# Function to save results to Excel with error handling
def save_results_to_excel(results, excel_path):
    df = pd.DataFrame(results, columns=['Certificate', 'QR Data', 'Name', 'Registration Number', 'URL Status'])
    
    try:
        df.to_excel(excel_path, index=False)
        print(f"Results successfully saved to {excel_path}")
    except PermissionError:
        print(f"PermissionError: Unable to save the file. Please close '{excel_path}' if it is open and try again.")
    except Exception as e:
        print(f"An error occurred while saving the file: {e}")

# Folder containing certificates
certificates_folder = 'certificates/'

# Scan all certificates in the folder
scanned_data = scan_certificates_in_folder(certificates_folder)

# Prepare the data to be saved in Excel
results = []
for certificate, qr_data, name, registration_number in scanned_data:
    if qr_data != "No QR code found":
        valid = is_url_valid(qr_data)
        validity_status = "Valid" if valid else "Invalid"
        results.append((certificate, qr_data, name, registration_number, validity_status))
    else:
        results.append((certificate, qr_data, name, registration_number, "N/A"))

# Save the results to an Excel file
excel_file_path = 'scanned_results_new.xlsx'  # Use a different file name to avoid conflicts
save_results_to_excel(results, excel_file_path)








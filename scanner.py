# import cv2
# from pyzbar.pyzbar import decode
# import os
# import requests

# def scan_qr_code(image_path):
#     img = cv2.imread(image_path) 
#     qr_codes = decode(img)  
    
#     if qr_codes:  
#         for qr_code in qr_codes:  
#             return qr_code.data.decode('utf-8') 
#     return None  

# def scan_certificates_in_folder(folder_path):
#     scanned_results = {} 
    
#     for certificate_file in os.listdir(folder_path):  
#         certificate_path = os.path.join(folder_path, certificate_file)  
        
#         if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):  
#             qr_data = scan_qr_code(certificate_path)  
            
#             if qr_data:  
#                 scanned_results[certificate_file] = qr_data  
#             else:  
#                 scanned_results[certificate_file] = "No QR code found"
    
#     return scanned_results  


# def is_url_valid(url):
#     try:
#         response = requests.head(url, allow_redirects=True, timeout=5)
#         return response.status_code == 200
#     except requests.RequestException:
#         return False

# # Folder containing certificates
# certificates_folder = 'certificates/'

# # Scan all certificates in the folder
# scanned_data = scan_certificates_in_folder(certificates_folder)

# # Output the scanned results and validate URLs
# for certificate, qr_data in scanned_data.items():  
#     if qr_data != "No QR code found":
#         valid = is_url_valid(qr_data)
#         validity_status = "Valid" if valid else "Invalid"
#         print(f"Certificate: {certificate}, QR Data: {qr_data}, URL Status: {validity_status}")
#     else:
#         print(f"Certificate: {certificate}, QR Data: {qr_data}")























# import cv2
# from pyzbar.pyzbar import decode
# import pytesseract
# import os
# import requests

# # Function to scan QR code
# def scan_qr_code(image_path):
#     img = cv2.imread(image_path)
#     qr_codes = decode(img)
    
#     if qr_codes:
#         for qr_code in qr_codes:
#             return qr_code.data.decode('utf-8')
#     return None

# # Function to extract text (name and registration number) using pytesseract
# def extract_name_and_registration(image_path):
#     img = cv2.imread(image_path)
#     text = pytesseract.image_to_string(img)
    
#     # Find the name and registration number in the extracted text
#     # Assuming the text has specific patterns for name and registration number
#     name = None
#     registration_number = None
    
#     for line in text.splitlines():
#         if "Name" in line:  # Adjust this according to the actual format of the scorecard
#             name = line.split(":")[-1].strip()
#         if "Registration" in line:  # Adjust this according to the actual format
#             registration_number = line.split(":")[-1].strip()
    
#     return name, registration_number

# # Function to scan certificates in the folder and get the QR data, name, and registration number
# def scan_certificates_in_folder(folder_path):
#     scanned_results = {}
    
#     for certificate_file in os.listdir(folder_path):
#         certificate_path = os.path.join(folder_path, certificate_file)
        
#         if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
#             qr_data = scan_qr_code(certificate_path)
#             name, registration_number = extract_name_and_registration(certificate_path)
            
#             if qr_data:
#                 scanned_results[certificate_file] = {
#                     "QR Data": qr_data,
#                     "Name": name or "Name not found",
#                     "Registration Number": registration_number or "Registration number not found"
#                 }
#             else:
#                 scanned_results[certificate_file] = {
#                     "QR Data": "No QR code found",
#                     "Name": name or "Name not found",
#                     "Registration Number": registration_number or "Registration number not found"
#                 }
    
#     return scanned_results

# # Function to validate URLs
# def is_url_valid(url):
#     try:
#         response = requests.head(url, allow_redirects=True, timeout=5)
#         return response.status_code == 200
#     except requests.RequestException:
#         return False

# # Folder containing certificates
# certificates_folder = 'certificates/'

# # Scan all certificates in the folder
# scanned_data = scan_certificates_in_folder(certificates_folder)

# # Output the scanned results and validate URLs
# for certificate, data in scanned_data.items():
#     qr_data = data["QR Data"]
    
#     if qr_data != "No QR code found":
#         valid = is_url_valid(qr_data)
#         validity_status = "Valid" if valid else "Invalid"
#         print(f"Certificate: {certificate}, Name: {data['Name']}, Registration Number: {data['Registration Number']}, QR Data: {qr_data}, URL Status: {validity_status}")
#     else:
#         print(f"Certificate: {certificate}, Name: {data['Name']}, Registration Number: {data['Registration Number']}, QR Data: {qr_data}")



































# import cv2
# from pyzbar.pyzbar import decode
# import pytesseract
# import os
# import requests

# # Set the path to tesseract executable if it's not in PATH
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# # Function to scan QR code
# def scan_qr_code(image_path):
#     img = cv2.imread(image_path)
#     qr_codes = decode(img)
    
#     if qr_codes:
#         for qr_code in qr_codes:
#             return qr_code.data.decode('utf-8')
#     return None

# # Function to extract text (name and registration number) using pytesseract
# def extract_name_and_registration(image_path):
#     img = cv2.imread(image_path)
#     text = pytesseract.image_to_string(img)
    
#     # Find the name and registration number in the extracted text
#     name = None
#     registration_number = None
    
#     for line in text.splitlines():
#         if "Name" in line:  # Adjust this according to the actual format of the scorecard
#             name = line.split(":")[-1].strip()
#         if "Registration" in line:  # Adjust this according to the actual format
#             registration_number = line.split(":")[-1].strip()
    
#     return name, registration_number

# # Function to scan certificates in the folder and get the QR data, name, and registration number
# def scan_certificates_in_folder(folder_path):
#     scanned_results = {}
    
#     for certificate_file in os.listdir(folder_path):
#         certificate_path = os.path.join(folder_path, certificate_file)
        
#         if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
#             qr_data = scan_qr_code(certificate_path)
#             name, registration_number = extract_name_and_registration(certificate_path)
            
#             if qr_data:
#                 scanned_results[certificate_file] = {
#                     "QR Data": qr_data,
#                     "Name": name or "Name not found",
#                     "Registration Number": registration_number or "Registration number not found"
#                 }
#             else:
#                 scanned_results[certificate_file] = {
#                     "QR Data": "No QR code found",
#                     "Name": name or "Name not found",
#                     "Registration Number": registration_number or "Registration number not found"
#                 }
    
#     return scanned_results

# # Function to validate URLs
# def is_url_valid(url):
#     try:
#         response = requests.head(url, allow_redirects=True, timeout=5)
#         return response.status_code == 200
#     except requests.RequestException:
#         return False

# # Folder containing certificates
# certificates_folder = 'certificates/'

# # Scan all certificates in the folder
# scanned_data = scan_certificates_in_folder(certificates_folder)

# # Output the scanned results and validate URLs
# for certificate, data in scanned_data.items():
#     qr_data = data["QR Data"]
    
#     if qr_data != "No QR code found":
#         valid = is_url_valid(qr_data)
#         validity_status = "Valid" if valid else "Invalid"
#         print(f"Certificate: {certificate}, Name: {data['Name']}, Registration Number: {data['Registration Number']}, QR Data: {qr_data}, URL Status: {validity_status}")
#     else:
#         print(f"Certificate: {certificate}, Name: {data['Name']}, Registration Number: {data['Registration Number']}, QR Data: {qr_data}")

































# import cv2
# from pyzbar.pyzbar import decode
# import pytesseract
# import os
# import re
# import requests

# # Set the path to tesseract executable if it's not in PATH
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# # Function to scan QR code
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
    
#     # Use regex to find name and registration number
#     name = None
#     registration_number = None

#     # Regex to capture name (assuming "Name:" precedes the name)
#     name_match = re.search(r"Name\s*:\s*(.*)", text)
#     if name_match:
#         name = name_match.group(1).strip()

#     # Regex to capture registration number (assuming "Registration No:" precedes the number)
#     reg_num_match = re.search(r"Registration(?:\s*No\s*|\s*Number\s*):\s*(\w+)", text)
#     if reg_num_match:
#         registration_number = reg_num_match.group(1).strip()
    
#     return name, registration_number

# # Function to scan certificates in the folder and get the QR data, name, and registration number
# def scan_certificates_in_folder(folder_path):
#     scanned_results = {}
    
#     for certificate_file in os.listdir(folder_path):
#         certificate_path = os.path.join(folder_path, certificate_file)
        
#         if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):  # Adjust if there are other formats
#             qr_data = scan_qr_code(certificate_path)
#             name, registration_number = extract_name_and_registration(certificate_path)
            
#             scanned_results[certificate_file] = {
#                 "QR Data": qr_data if qr_data else "No QR code found",
#                 "Name": name or "Name not found",
#                 "Registration Number": registration_number or "Registration number not found"
#             }
    
#     return scanned_results

# # Function to validate URLs
# def is_url_valid(url):
#     try:
#         response = requests.head(url, allow_redirects=True, timeout=5)
#         return response.status_code == 200
#     except requests.RequestException:
#         return False

# # Folder containing certificates
# certificates_folder = 'certificates/'

# # Scan all certificates in the folder
# scanned_data = scan_certificates_in_folder(certificates_folder)

# # Output the scanned results and validate URLs
# for certificate, data in scanned_data.items():
#     qr_data = data["QR Data"]
    
#     if qr_data != "No QR code found":
#         valid = is_url_valid(qr_data)
#         validity_status = "Valid" if valid else "Invalid"
#         print(f"Certificate: {certificate}, Name: {data['Name']}, Registration Number: {data['Registration Number']}, QR Data: {qr_data}, URL Status: {validity_status}")
#     else:
#         print(f"Certificate: {certificate}, Name: {data['Name']}, Registration Number: {data['Registration Number']}, QR Data: {qr_data}")









































import cv2
from pyzbar.pyzbar import decode
import pytesseract
import os
import re
import requests

# Set the path to tesseract executable if it's not in PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to scan all QR codes
def scan_qr_codes(image_path):
    img = cv2.imread(image_path)
    qr_codes = decode(img)
    
    qr_data_list = []
    if qr_codes:
        for qr_code in qr_codes:
            qr_data_list.append(qr_code.data.decode('utf-8'))
    
    return qr_data_list if qr_data_list else None

# Function to extract name and registration number using pytesseract
def extract_name_and_registration(image_path):
    img = cv2.imread(image_path)
    text = pytesseract.image_to_string(img)
    
    # Use regex to find name and registration number
    name = None
    registration_number = None

    # Regex to capture name (assuming "Name:" precedes the name)
    name_match = re.search(r"Name\s*:\s*(.*)", text)
    if name_match:
        name = name_match.group(1).strip()

    # Regex to capture registration number (assuming "Registration No:" precedes the number)
    reg_num_match = re.search(r"Registration(?:\s*No\s*|\s*Number\s*):\s*(\w+)", text)
    if reg_num_match:
        registration_number = reg_num_match.group(1).strip()
    
    return name, registration_number

# Function to scan certificates in the folder and get the QR data, name, and registration number
def scan_certificates_in_folder(folder_path):
    scanned_results = {}
    
    for certificate_file in os.listdir(folder_path):
        certificate_path = os.path.join(folder_path, certificate_file)
        
        if certificate_file.endswith(('.png', '.jpg', '.jpeg', '.pdf')):  # Adjust if there are other formats
            qr_data_list = scan_qr_codes(certificate_path)
            name, registration_number = extract_name_and_registration(certificate_path)
            
            scanned_results[certificate_file] = {
                "QR Data": qr_data_list if qr_data_list else ["No QR code found"],
                "Name": name or "Name not found",
                "Registration Number": registration_number or "Registration number not found"
            }
    
    return scanned_results

# Function to validate URLs
def is_url_valid(url):
    try:
        response = requests.head(url, allow_redirects=True, timeout=5)
        return response.status_code == 200
    except requests.RequestException:
        return False

# Folder containing certificates
certificates_folder = 'certificates/'

# Scan all certificates in the folder
scanned_data = scan_certificates_in_folder(certificates_folder)

# Output the scanned results and validate URLs
for certificate, data in scanned_data.items():
    print(f"\nCertificate: {certificate}")
    print(f"Name: {data['Name']}")
    print(f"Registration Number: {data['Registration Number']}")
    
    qr_data_list = data["QR Data"]
    for qr_data in qr_data_list:
        if qr_data != "No QR code found":
            valid = is_url_valid(qr_data)
            validity_status = "Valid" if valid else "Invalid"
            print(f"QR Data: {qr_data}, URL Status: {validity_status}")
        else:
            print("QR Data: No QR code found")

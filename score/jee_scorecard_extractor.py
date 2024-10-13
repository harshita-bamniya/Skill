import os
import re
import pandas as pd
import pytesseract
from PyPDF2 import PdfReader
from PIL import Image

def extract_details_from_pdf(pdf_file):
    try:
        reader = PdfReader(pdf_file)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

        print(f"Extracted text from {pdf_file}:\n{text}\n")

        # Adjust regex patterns to capture required details from PDFs
        reg_no = re.search(r"[A-Z0-9]{13}", text)  # Adjust based on your registration number format
        name = re.search(r"Name of the Candidate\s*:\s*([A-Z\s]+)", text)  # Extract name
        gate_score = re.search(r"GATE Score\s*:\s*(\d+)", text)  # Extract GATE score

        reg_no = reg_no.group(0).strip() if reg_no else None
        name = name.group(1).strip() if name else None
        gate_score = gate_score.group(1).strip() if gate_score else None

        return [name, gate_score, reg_no]
    except Exception as e:
        print(f"Error extracting details from {pdf_file}: {e}")
        return None

def extract_details_from_image(image_file):
    try:
        img = Image.open(image_file)

        # Optional: Convert image to grayscale to enhance text recognition
        img = img.convert('L')  # Convert to grayscale
        img = img.point(lambda x: 0 if x < 128 else 255, '1')  # Binarize image (black & white)

        # Use Tesseract to extract text
        text = pytesseract.image_to_string(img, config='--psm 6')  # Assume a single uniform block of text
        print(f"Extracted text from {image_file}:\n{text}\n")

        # Adjust regex patterns to capture required details from images
        reg_no = re.search(r"[A-Z0-9]{13}", text)  # Adjust based on your registration number format
        name = re.search(r"Name\s*:\s*([A-Z\s]+)", text)  # Extract name
        gate_score = re.search(r"GATE Score\s*:\s*(\d+)", text)  # Extract GATE score

        reg_no = reg_no.group(0).strip() if reg_no else None
        name = name.group(1).strip() if name else None
        gate_score = gate_score.group(1).strip() if gate_score else None

        return [name, gate_score, reg_no]
    except Exception as e:
        print(f"Error extracting details from {image_file}: {e}")
        return None

def extract_details(file_path):
    if file_path.lower().endswith('.pdf'):
        return extract_details_from_pdf(file_path)
    elif file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
        return extract_details_from_image(file_path)
    else:
        print(f"Unsupported file format: {file_path}")
        return None

def save_to_excel(data, filename):
    if data:
        df = pd.DataFrame(data, columns=["Name", "GATE Score", "Registration Number"])  # Adjusted column headers
        print(f"Saving data to Excel: {df}")  # Debugging statement
        try:
            df.to_excel(filename, index=False)
            print(f"Data successfully saved to {filename}")
        except Exception as e:
            print(f"Error saving to Excel: {e}")
    else:
        print("No data to save.")

def main(folder_path):
    all_data = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        print(f"Processing file: {file_path}")  # Debugging statement
        details = extract_details(file_path)
        if details and None not in details:
            print(f"Extracted details: {details}")  # Debugging statement
            all_data.append(details)
        else:
            print(f"Incomplete or invalid details, skipping: {details}")

    save_to_excel(all_data, 'jee_scorecard_data.xlsx')

if __name__ == "__main__":
    folder_path = input("Enter the folder path containing the JEE scorecards: ")
    main(folder_path)

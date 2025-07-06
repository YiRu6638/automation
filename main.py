import fitz
import os
import pandas as pd
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
from pytesseract import image_to_string
from PIL import Image
import pytesseract
from pathlib import Path
import sys
import layoutparser as lp
import cv2
import re
import matching

def mass_convert_pdf_to_img(pdf_directory, img_directory):
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define input and output folders
    pdf_folder = os.path.join(script_dir, pdf_directory)
    output_folder = os.path.join(script_dir, img_directory)
    
    os.makedirs(output_folder, exist_ok=True)

    for file_name in os.listdir(pdf_folder):
    # Process only PDF files
        if file_name.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, file_name)

            try:
                # Convert the first page of the PDF to an image
                images = convert_from_path(pdf_path)
                
                # Extract the base name (without extension)
                base_name = os.path.splitext(file_name)[0]

                # Save the first page as an image in the dummy_img folder
                output_image_path = os.path.join(output_folder, f"{base_name}.jpg")
                images[0].save(output_image_path, "JPEG")

                print(f"Saved image for {file_name} as {base_name}.jpg in dummy_img")
            except Exception as e:
                print(f"Error processing {file_name}: {e}")

def mass_convert_img_to_text(img_directory, text_directory):
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Define input and output folders
    image_folder = os.path.join(script_dir, img_directory)
    text_folder = os.path.join(script_dir, text_directory)
    
    os.makedirs(text_folder, exist_ok=True)
    
    for image_name in os.listdir(image_folder):
        if image_name.endswith(".jpg"):
            image_path = os.path.join(image_folder, image_name)
            try:
                # Extract text using Tesseract OCR
                text = image_to_string(image_path)
                base_name = os.path.splitext(image_name)[0]
                output_text_path = os.path.join(text_folder, f"{base_name}.txt")
                with open(output_text_path, "w", encoding="utf-8") as text_file:
                    text_file.write(text)
                print(f"Saved text for {image_name} as {base_name}.txt in dummy_text")
            except Exception as e:
                print(f"Error processing {image_name}: {e}")

# Not used
def alternative_pdf_extraction_method():    
    img = cv2.imread(single_img)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    text = image_to_string(gray, lang='eng')  # Specify language if necessary

    # Ensure the output directory exists
    os.makedirs(os.path.dirname(output_text_file), exist_ok=True)

    # Save extracted text to a file
    with open(output_text_file, "w", encoding="utf-8") as file:
        file.write(text)

    print(f"Text has been saved to {output_text_file}")

if __name__ == "__main__":
    mass_convert_pdf_to_img("dummy_pdf", "dummy_img")
    mass_convert_img_to_text("dummy_img", "dummy_text")
    # matching.extract_text_to_excel("dummy_text")
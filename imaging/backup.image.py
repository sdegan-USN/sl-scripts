import cv2
import numpy as np
from pdf2image import convert_from_path
from openpyxl import Workbook
import pytesseract

# Make sure to install pytesseract and tesseract-ocr
# pip install pytesseract
# For Ubuntu: sudo apt-get install tesseract-ocr
# For MacOS: brew install tesseract

# Function to convert a PDF file into a list of images
def pdf_to_images(pdf_file):
    return convert_from_path(pdf_file)

# Function to find table contours in an image
def find_tables(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    table_contours = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w > 100 and h > 100:
            table_contours.append((x, y, w, h))

    return table_contours

# Function to extract cell contours from a table
def extract_table(image, table_contour):
    x, y, w, h = table_contour
    table = image[y:y + h, x:x + w]
    gray = cv2.cvtColor(table, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]

    contours, _ = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    cell_contours = [cv2.boundingRect(c) for c in contours if cv2.contourArea(c) > 50]

    return cell_contours

# Function to read text from a cell and return a list of text-color pairs
def read_text_from_cell(image, cell_contour):
    x, y, w, h = cell_contour
    cell = image[y:y + h, x:x + w]

    # Define a list of colors (white, red, and black) to extract the text
    colors = [(255, 255, 255), (0, 0, 255), (0, 0, 0)]
    # Initialize an empty list to store the extracted text-color pairs
    text_color_pairs = []
    # Iterate through the colors
    for color in colors:
        lower = np.array(color, dtype="uint8")
        upper = np.array(color, dtype="uint8")
        mask = cv2.inRange(cell, lower, upper)
        text = pytesseract.image_to_string(mask, config="--psm 6").strip()
        if text:
            text_color_pairs.append((text, color))

    return text_color_pairs

# Function to process a PDF file and save the extracted tables to an Excel file
def process_pdf(pdf_file, output_file):
    images = pdf_to_images(pdf_file)
    workbook = Workbook()
    sheet = workbook.active

    for i, image in enumerate(images):
        print("Processing image "+str(i))
        image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        table_contours = find_tables(image)

        for table_contour in table_contours:
            print("Processing table "+str(table_contour))
            cell_contours = extract_table(image, table_contour)
            data = []

            for cell_contour in cell_contours:
                print("Processing contour "+str(cell_contour))
                text_color_pairs = read_text_from_cell(image, cell_contour)
                print("Text color processing")
                cell_data = " ".join([text for text, _ in text_color_pairs])
                print("Cell data "+str(cell_data))
                data.append((cell_contour[1], cell_data))

            data.sort(key=lambda x: x[0])

            for j, (_, cell_data) in enumerate(data):
                sheet.cell(row=j + 1, column=i + 1, value=cell_data)

    workbook.save(output_file)


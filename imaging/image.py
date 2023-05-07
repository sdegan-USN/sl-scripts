import cv2
import numpy as np
from pdf2image import convert_from_path
from openpyxl import Workbook
import pytesseract
from multiprocessing import Pool, cpu_count
import random

# ... (keep the existing functions pdf_to_images, find_tables, extract_table, and read_text_from_cell)

# Function to convert a PDF file into a list of images
def pdf_to_images(pdf_file):
    return convert_from_path(pdf_file)

# Function to extract cell contours from a table
def extract_table(image, table_contour):
    x, y, w, h = table_contour
    table = image[y:y + h, x:x + w]
    gray = cv2.cvtColor(table, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]

    contours, _ = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    cell_contours = [cv2.boundingRect(c) for c in contours if cv2.contourArea(c) > 50]

    return cell_contours


# Function to check if an image contains a cell with a blue background and white text
def has_blue_bg_white_text(image):
    print("has_blue_bg_white called")
    lower_blue_bg = np.array([40, 100, 190], dtype="uint8")
    upper_blue_bg = np.array([60, 130, 210], dtype="uint8")

    lower_white_text = np.array([240, 240, 240], dtype="uint8")
    upper_white_text = np.array([255, 255, 255], dtype="uint8")

    blue_mask = cv2.inRange(image, lower_blue_bg, upper_blue_bg)
    white_mask = cv2.inRange(image, lower_white_text, upper_white_text)

    blue_white_mask = cv2.bitwise_and(blue_mask, white_mask)

    print("has_blue_bg_white finished")
    return cv2.countNonZero(blue_white_mask) > 0


# Update the find_tables function
def find_tables(image):
    print("find_tables called")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    random_num = random.randint(1,1000)
    cv2.imwrite(f"thresh_{random_num}.jpg", thresh)  # Save the thresholded image with a unique filename
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    table_contours = []
    for contour in contours:
        #print(contour)
        print("finding contours")
        x, y, w, h = cv2.boundingRect(contour)
        if w > 100 and h > 100:
            print(x,y,w,h)
            table = image[y:y + h, x:x + w]
            #if has_blue_bg_white_text(table):
            #    table_contours.append((x, y, w, h))
            table_contours.append((x, y, w, h))
        print("contours found")
    print("find_tables finished")
    return table_contours

# Update the read_text_from_cell function
def read_text_from_cell(image, cell_contour):
    x, y, w, h = cell_contour
    cell = image[y:y + h, x:x + w]

    # Define a list of colors (white and black) to extract the text
    colors = [(255, 255, 255), (0, 0, 0)]
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

# Function to process a single image and return the extracted data
def process_image(image):
    data = []
    image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    table_contours = find_tables(image)
    print(" ")
    print("FIND TABLES HAS COMPLETED")

    for table_contour in table_contours:
        print("extract_table called")
        cell_contours = extract_table(image, table_contour)
        print("extract_table finished")
        
        print("looping over cell contours")
        for cell_contour in cell_contours:
            #print("read_text_from_cell called")
            text_color_pairs = read_text_from_cell(image, cell_contour)
            #print("read_text_from_cell finished")
            cell_data = " ".join([text for text, _ in text_color_pairs])
            data.append((cell_contour[1], cell_data))
        print("finished looping over cell contours")
    return data

# Function to process a PDF file and save the extracted tables to an Excel file
def process_pdf(pdf_file, output_file):
    images = pdf_to_images(pdf_file)
    workbook = Workbook()
    sheet = workbook.active

    with Pool(cpu_count()) as pool:
        results = pool.map(process_image, images)

    for i, data in enumerate(results):
        data.sort(key=lambda x: x[0])

        for j, (_, cell_data) in enumerate(data):
            sheet.cell(row=j + 1, column=i + 1, value=cell_data)

    workbook.save(output_file)


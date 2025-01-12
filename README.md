# py_Photo2Excel_data_convert
python to identify photo of Excel-like paper and get data, input the data to Spread Sheets(xlsx)  within the format

To automate the process of extracting numerical data from a photo of an Excel-like report and writing it into a Numbers spreadsheet on macOS, 

1. Image Preprocessing: Convert the image to a format suitable for text extraction (such as grayscale and thresholding).
2. Optical Character Recognition (OCR): Extract the numbers (and possibly headers) from the image.
3. Data Parsing: Structure the extracted text into a tabular format.
4. Write Data to Numbers Spreadsheet: Use AppleScript (through Python) to write the parsed data into a Numbers file.

Here’s how you can implement this workflow using Python.

## Step 1: Install Required Libraries
You'll need several libraries to perform image processing, OCR, and automation tasks. Install the following Python libraries:
```
pip install pytesseract opencv-python numpy pillow
```

pytesseract is the Python wrapper for Tesseract OCR, which will be used to extract text from the image.
opencv-python will help in preprocessing the image (grayscale, thresholding).
pillow is for image manipulation (if needed).
Additionally, you'll need Tesseract OCR installed on your machine. You can install it via Homebrew on macOS:
```
brew install tesseract
```

## Step 2: Preprocess the Image and Extract Text Using OCR
Let's start by preprocessing the image to improve the OCR results and then extract text (numbers) from the image.

```
import cv2
import pytesseract
import numpy as np
from PIL import Image

# Set the path for Tesseract (if not set in the environment variable)
pytesseract.pytesseract.tesseract_cmd = '/usr/local/bin/tesseract'  # Adjust path as needed

def extract_numbers_from_image(image_path):
    # Load the image
    img = cv2.imread(image_path)

    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Apply thresholding to binarize the image (improves OCR accuracy)
    _, thresh_img = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    # Use Tesseract to extract text (numbers)
    custom_oem_psm_config = r'--oem 3 --psm 6'  # OEM = 3 (default), PSM = 6 (Assume a single uniform block of text)
    text = pytesseract.image_to_string(thresh_img, config=custom_oem_psm_config)

    # For debugging: Print extracted text
    print("Extracted Text: \n", text)

    return text

# Example usage
image_path = "path_to_your_excel_like_report_image.jpg"  # Replace with the path to your image
extracted_text = extract_numbers_from_image(image_path)

```
Explanation:

Grayscale Conversion: The image is converted to grayscale to simplify the image and improve OCR accuracy.
Thresholding: A binary thresholding technique is applied to make the numbers stand out from the background.
Tesseract OCR: This is used to extract the text (numbers) from the image.


## Step 3: Parse Extracted Data
Once we extract the text, we need to clean it up and structure it into a tabular format (assuming the report follows a grid-like structure).

```
import re

def parse_extracted_text_to_table(extracted_text):
    # Split by newlines and remove any non-numeric characters
    lines = extracted_text.split("\n")
    data = []

    for line in lines:
        # Remove non-numeric characters except for commas and dots
        cleaned_line = re.sub(r'[^0-9,\.]', '', line)
        
        # If the cleaned line contains any numbers, split by whitespace
        if cleaned_line.strip():
            row = cleaned_line.split()
            data.append(row)

    return data

# Example usage
parsed_data = parse_extracted_text_to_table(extracted_text)
print("Parsed Data: \n", parsed_data)

```
Explanation:

Regex Cleaning: We use regular expressions to clean the extracted text by keeping only the numeric values (numbers, dots, and commas).
Data Splitting: The data is split into rows, and each row is split into individual cells by whitespace.


## Step 4: Write Data to Numbers Spreadsheet
Now that we have the data in a structured format (a list of rows and columns), we can write it to a Numbers spreadsheet. To do this, we’ll use AppleScript called from Python to interact with Numbers.

```
import subprocess

def write_data_to_numbers(parsed_data):
    # AppleScript to write data to a Numbers file
    applescript = """
    tell application "Numbers"
        activate
        set doc to open "Macintosh HD:Users:YourUsername:Documents:example.numbers"  -- Adjust path as needed
        tell sheet 1 of doc
    """
    
    # Loop through the rows of parsed data and write them to Numbers
    for i, row in enumerate(parsed_data):
        for j, cell_value in enumerate(row):
            applescript += f"""
                set value of cell {i + 1}, {j + 1} of sheet 1 of doc to "{cell_value}"
            """
    
    # Finish the AppleScript to save and close the document
    applescript += """
        save doc
        close doc
    end tell
    """

    # Run the AppleScript using osascript
    process = subprocess.Popen(['osascript', '-e', applescript], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()

    if stderr:
        print(f"Error: {stderr.decode()}")
    else:
        print("Successfully wrote data to Numbers sheet.")

# Example usage
write_data_to_numbers(parsed_data)

```

Explanation:

AppleScript: The script opens an existing Numbers document, and then for each cell in the parsed data, it sets the value in the corresponding cell in the Numbers spreadsheet.
Subprocess: The Python subprocess module is used to execute the AppleScript, which interacts with the Numbers app.
Putting Everything Together


# Here’s how to tie everything together in a single script:

```
import cv2
import pytesseract
import numpy as np
import re
import subprocess

# Set Tesseract path if not set in environment variables
pytesseract.pytesseract.tesseract_cmd = '/usr/local/bin/tesseract'  # Adjust path as needed

def extract_numbers_from_image(image_path):
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh_img = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
    
    custom_oem_psm_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(thresh_img, config=custom_oem_psm_config)
    
    return text

def parse_extracted_text_to_table(extracted_text):
    lines = extracted_text.split("\n")
    data = []

    for line in lines:
        cleaned_line = re.sub(r'[^0-9,\.]', '', line)
        if cleaned_line.strip():
            row = cleaned_line.split()
            data.append(row)

    return data

def write_data_to_numbers(parsed_data):
    applescript = """
    tell application "Numbers"
        activate
        set doc to open "Macintosh HD:Users:YourUsername:Documents:example.numbers"
        tell sheet 1 of doc
    """
    
    for i, row in enumerate(parsed_data):
        for j, cell_value in enumerate(row):
            applescript += f"""
                set value of cell {i + 1}, {j + 1} of sheet 1 of doc to "{cell_value}"
            """
    
    applescript += """
        save doc
        close doc
    end tell
    """
    
    process = subprocess.Popen(['osascript', '-e', applescript], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()

    if stderr:
        print(f"Error: {stderr.decode()}")
    else:
        print("Successfully wrote data to Numbers sheet.")

# Main execution
image_path = "path_to_your_excel_like_report_image.jpg"
extracted_text = extract_numbers_from_image(image_path)
parsed_data = parse_extracted_text_to_table(extracted_text)
write_data_to_numbers(parsed_data)

```
Explanation:
1. extract_numbers_from_image: Preprocesses the image and extracts text using OCR.
2. parse_extracted_text_to_table: Cleans the extracted text and structures it into a table.
3. write_data_to_numbers: Uses AppleScript to write the parsed data into a Numbers spreadsheet.

   
Conclusion
This solution uses OCR to extract data from an image, structures the data into a tabular format, and writes the data into an existing Numbers spreadsheet using AppleScript. Adjust the file paths and tweak the preprocessing steps (e.g., thresholding) for your specific case.


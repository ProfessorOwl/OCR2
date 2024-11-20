import cv2
import numpy as np
import pytesseract
import re
import pandas as pd
import openpyxl
from os import listdir
from os.path import isfile, join

# set correct path for pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\czirbes\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

# load image
def path_to_image(number, path):
    img = cv2.imread(f"{path}/Volume_{number}.jpg")
    return img

# base function that returns text from the image
def ocr_core(img):
    text = pytesseract.image_to_string(img)
    return text

# function that automatically detects the image numbers "from" and "to"

def file_from_to(path):
    filenames = [f for f in listdir(path) if isfile(join(path, f))]
    filenames_cleared = [name.split("_")[1] for name in filenames if name.split("_")[0] == "Volume"]
    filenames = None
    filenames_cleared = None
    return filenames_cleared[0].split(".")[0], filenames_cleared[-1].split(".")[0]

# function that is tailored to the output of the DoD-printer software. 
def mult_ocr(from_number, to_number, filename, sheetname="data", path="Bilder/"):
    """Iterates through multiple files and prints the ocr output into an excel file.
    
    positional arguments:
    from_number -- from which number to start counting the files
    to_number -- to which number to count the files
    filename -- name of excel file to be created
    sheetname -- name of excel sheet in the file
    path -- path where the images are stored 
    """
    
    list = [] 
    for number in range(from_number, to_number+1):
        
        # ---- Prints the percentage of completed pictures ----
        num_of_images = to_number+1 - from_number
        progress = round(((number-from_number)/num_of_images)*100, 1) # Rounds to the first digit after comma
        print(f"Progress: {progress}%") # Print the percentage
        LINE_UP = '\033[1A' # ASCII code to go to the previous line
        LINE_CLEAR = '\x1b[2K' # ASCII code to clear the current line
        print(LINE_UP, end=LINE_CLEAR) # GO a line up and clear after printing
        # --------
        
        # ---- Do some image processing
        img = path_to_image(number, path)[0:40, 10:125] # Get image and crop it to the desired size to reduce computing time
        text = ocr_core(img) # Get the text via OCR
        while True:
            try:
                value = float(".".join(re.findall("\d+", text)))
                break
            except ValueError:
                print(f"I just skipped picture {number}, because I couldn't convert \"{value}\" into a number.")
                value = None
                break
        converted = [number, value] # Create list and put image in the first slot. Second slot 
                                                                       # is the OCR'ed text, which is scanned for all digits. Since both digits
                                                                       # are separated by a dot, this has to be added via the "join" function
        list.append(converted) # Add the converted element to our list 
    df = pd.DataFrame(list) # Convert the list into a dataframe for export to excel
    df.columns=["filename", "volume / nL"] # Name the columns
    with pd.ExcelWriter(f"{filename}.xlsx", mode="a") as writer: # Uses the "write" function, so that a new sheet can be added to an existing excel file        
          df.to_excel(writer, sheet_name=f"{sheetname}") # Defines, what the sheet should be named as
    print(f"Converted {abs(from_number-to_number)} images.") # Prints how many images where converted
    return


#mult_ocr(6624, # First number
#        7565, # Last number
#        "a_PJ200 L", # Name of Excel file
#        "n_Stroke 0 Velocity 85 1Hz", # Name of Sheet in Excel file, no special symbols allowed!
#        "Bilder/a_PJ200 L/n_Stroke 0 Velocity 85 1Hz 2,46ms") # Path to the folder with all the pictures

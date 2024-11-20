import cv2
import numpy as np
import pytesseract
import re
import pandas as pd
import openpyxl
from os import listdir
from os.path import isfile, join, exists

# set correct path for pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\czirbes\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

# load image
def path_to_image(number, path):
    number = "{:06d}".format(number)
    img = cv2.imread(f"{path}/Volume_{number}.jpg")
    return img

# base function that returns text from the image
def ocr_core(img):
    text = pytesseract.image_to_string(img)
    return text

# function that automatically detects the image numbers "from" and "to"

def file_from_to(path):
    filenames_cleared = None
    filenames = [f for f in listdir(path) if isfile(join(path, f))] # get every name of file-types in the folder
    filenames_cleared = [name.split("_")[1] for name in filenames if name.split("_")[0] == "Volume"] # Select only the number if the file is named "Volume"
    filenames = None # clear filenames to save space
    return filenames_cleared[0].split(".")[0], filenames_cleared[-1].split(".")[0] # return the first and last filename


# function that is tailored to the output of the DoD-printer software. 
def mult_ocr(from_number, 
             to_number, 
             path, 
             sheetname="data", 
             imagepath="Bilder/", 
             sheet_exists_toggle="new", 
             progressbar = None,
             console = None):
    """Iterates through multiple files and prints the ocr output into an excel file.
    
    positional arguments:
    from_number -- from which number to start counting the files
    to_number -- to which number to count the files
    path -- name and path of excel file to be created
    sheetname -- name of excel sheet in the file
    imagepath -- path where the images are stored
    sheet_exists_toggle -- Sets which object should switch the overwrite mode of existing sheets. Standard: "new"
    """
    
    list = [] 
    for number in range(from_number, to_number+1):
        
        # ---- Prints the percentage of completed pictures ----
        num_of_images = to_number+1 - from_number
        progress = (number-from_number)/num_of_images # Rounds to the first digit after comma
        if progressbar == None:
            print(f"Progress: {progress}%") # Print the percentage
            LINE_UP = '\033[1A' # ASCII code to go to the previous line
            LINE_CLEAR = '\x1b[2K' # ASCII code to clear the current line
            print(LINE_UP, end=LINE_CLEAR) # Go a line up and clear after printing
        else:
            progressbar.set(progress)
            progressbar.update_idletasks()
        # --------
        
        # ---- Do some image processing
        img = path_to_image(number, imagepath)[0:40, 10:125] # Get image and crop it to the desired size to reduce computing time
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
    if exists(path):
        with pd.ExcelWriter(path, mode="a", if_sheet_exists=sheet_exists_toggle) as writer: # Uses the "append" function, so that a new sheet can be added to an existing excel file        
            df.to_excel(writer, sheet_name=f"{sheetname}") # Defines, what the sheet should be named as
    else:
        with pd.ExcelWriter(path, mode="w") as writer: # Uses the "write" function, so that a file will be created
            df.to_excel(writer, sheet_name=f"{sheetname}") 
    if console == None:
        print(f"Converted {abs(from_number-to_number)} images.") # Prints how many images where converted
    else:
        console.insert("0.0", f"Converted {abs(from_number-to_number)} images to sheet {sheetname} in {path}.")
    return
import tkinter as tk
from tkinter.filedialog import askdirectory, asksaveasfilename
import customtkinter as ctk
import init as ocr

root = ctk.CTk()
root.geometry("600x300")
root.title("DoD Results OCR")

def select_folder():
    folder_path = askdirectory()
    if not folder_path == "": # Delete the entry line if "Ask Directory" window gets closed
        entry_selectfolder.delete(0,tk.END)
    entry_selectfolder.insert(0,folder_path)
    fromimage.set(ocr.file_from_to(folder_path)[0])
    toimage.set(ocr.file_from_to(folder_path)[1])

def save_where():
    file_path = asksaveasfilename(
        confirmoverwrite=False,
        defaultextension=".xlsx",
        filetypes=[("Excel 2007 and higher", "*.xlsx")]
        )
    if not file_path == "": # Delete the entry line if "Ask Directory" window gets closed
        entry_savefile.delete(0,tk.END)
    entry_savefile.insert(0,file_path)
    
def togglenumberinput():
    if switch_selectfolder.get() == 1:
        frame_imagenumbers.grid()
    else:
        frame_imagenumbers.grid_remove()

#---- Create Frame for folder path, save path and sheet name
frame_selectsave = ctk.CTkFrame(
    master=root
)
frame_selectsave.columnconfigure(1, weight=1)
frame_selectsave.columnconfigure(2, minsize=320)

#----

#---- Construct the "Select Folder" button together with the entry line and the image number selector
btn_selectfolder = ctk.CTkButton(
    master=frame_selectsave,
    text="Select Folder",
    command=select_folder
)
entry_selectfolder = ctk.CTkEntry(
    master=frame_selectsave,
)
switch_selectfolder = ctk.CTkSwitch(
    master=frame_selectsave,
    text="Specify image numbers",
    command=togglenumberinput
)
btn_selectfolder.grid(row=0, column=0, padx=5, pady=10)
entry_selectfolder.grid(row=0, column=1, sticky="ew", padx=5, pady=10)
switch_selectfolder.grid(row=0, column=2, padx=5, pady=10)
#----

#---- Construct the "Save As..." button together with the entry line
btn_savefile = ctk.CTkButton(
    master=frame_selectsave,
    text="Save As...",
    command=save_where
)
entry_savefile = ctk.CTkEntry(
    master=frame_selectsave,
)
btn_savefile.grid(row=1, column=0, padx=5, pady=10)
entry_savefile.grid(row=1, column=1, sticky="ew", padx=5, pady=10)
#----

#---- Construct the Image Number entry fields
fromimage = ctk.StringVar()
toimage = ctk.StringVar()
frame_imagenumbers = ctk.CTkFrame(
    master=frame_selectsave
)
entry_fromimagenumber = ctk.CTkEntry(
    master=frame_imagenumbers,
    width=100,
    textvariable=fromimage
)
label_betweennumbers = ctk.CTkLabel(
    master=frame_imagenumbers,
    text="to",
    width=20
)
entry_toimagenumber = ctk.CTkEntry(
    master=frame_imagenumbers,
    width=100,
    textvariable=toimage
)
frame_imagenumbers.grid(row=1, column=2, ipady=10, ipadx=20)
entry_fromimagenumber.grid(row=0, column=0)
label_betweennumbers.grid(row=0, column=1)
entry_toimagenumber.grid(row=0, column=2)
frame_imagenumbers.columnconfigure([0,1,2], weight=1)
frame_imagenumbers.rowconfigure(0, weight=1)
frame_imagenumbers.grid_remove() # Removes the Label from the view
#----

#---- Construct the "Sheet Name" label, the entry line and the overwrite check
label_sheet_name = ctk.CTkLabel(
    master=frame_selectsave,
    text="Sheet Name"
)
entry_sheet_name = ctk.CTkEntry(
    master=frame_selectsave,
    placeholder_text="Sheet Name"
)
switch_sheet_name = ctk.CTkSwitch(
    master=frame_selectsave,
    text="Overwrite existing sheet?",
    onvalue="replace",
    offvalue="new"
)
label_sheet_name.grid(row=2, column=0, padx=5, pady=10)
entry_sheet_name.grid(row=2, column=1, sticky="ew", padx=5, pady=10)
switch_sheet_name.grid(row=2, column=2, padx=5, pady=10)
#----

#---- Button to start conversion and progress bar
frame_convertbar = ctk.CTkFrame(
    master=root
)
btn_convert = ctk.CTkButton(
    master=frame_convertbar,
    text="Convert images",
    command=lambda: ocr.mult_ocr(int(fromimage.get()),
                                 int(toimage.get()), 
                                 entry_savefile.get(), 
                                 imagepath=entry_selectfolder.get(),
                                 sheetname=entry_sheet_name.get(),
                                 sheet_exists_toggle=switch_sheet_name.get(),
                                 progressbar=progbar,
                                 console=text_output)
)
progbar = ctk.CTkProgressBar(
    master=frame_convertbar
)
progbar.set(0)
btn_convert.grid(row=0, column=0, padx=5, pady=10)
progbar.grid(row=0, column=1, sticky="ew", padx=5, pady=10)
frame_convertbar.columnconfigure(1, weight=1)
#----

#---- Text output on the bottom of the page
text_output = ctk.CTkTextbox(
    master=root,
    state="disabled"
)
#----

#---- Pack everything up
frame_selectsave.pack(
    fill="x"
)
frame_convertbar.pack(
    fill="x"
)
text_output.pack(
    fill="x"
)
#----


root.mainloop()
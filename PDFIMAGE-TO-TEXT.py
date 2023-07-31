# -*- coding: utf-8 -*-
"""
Created on Wed Jul 12 15:55:13 2023
@author: Shristi
"""
import PySimpleGUI as sg
import fitz
import pytesseract
import tempfile
from PIL import Image
import io
import docx

tesrct_path = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = tesrct_path

# Function to get the selected language
def get_language(values):
    if values["-HIN-"]:
        return "hin"
    elif values["-ENG-"]:
        return "eng"
    else:
        return "hin+eng"
    
def extract_img(pdf_file,output_dir,str_page, end_page):
    doc = fitz.open(pdf_file)
    if end_page == 0:
        end_page = str_page
    img_list = []

    # Iterate through each page of the PDF
    for page_num in range(str_page - 1,end_page):
        # Get the current page object
        page = doc.load_page(page_num)

        # Render the current page as a pixmap
        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72), alpha=False)
        
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
            f.write(pix.tobytes())
            img_list.append(f.name)

    # Close the current PDF file to free up resources
    doc.close()
    return(img_list)

def extract_text(image, lang):
    try:
        # Convert the image to grayscale
        image = image.convert('L')
        # Use OCR to extract text from the image
        text = pytesseract.image_to_string(image, lang=lang)
        return text
    except:
        print("Error in extracting text")
        return ""
    
def img_disp(img_path):
        image = Image.open(img_path)
        width, height = image.size
        
        # reduce the size of image for display
        ratio = int(width/300)
        # Print the size
        width = int(width/ratio)
        height = int(height/ratio)
        # print(f"Image size: {width} x {height}")
        image1 = image.resize((width, height))
        
        # Convert the image to bytes
        with io.BytesIO() as bytes_buffer:
            image1.save(bytes_buffer, format='PNG')
            image_bytes = bytes_buffer.getvalue()
        
        window["-PDF_IMAGE-"].update(data=image_bytes)

def process_clipbrd(lang):
        from PIL import ImageGrab
        # Get the image from the clipboard
        image = ImageGrab.grabclipboard()
        # Check if an image was found in the clipboard
        if image:
            # Extract text from the image based on the selected language
            text = extract_text(image, lang)
            
            # Update the image and text on the form
            try:
                # Create a PySimpleGUI Image object from the byte buffer
                img_buffer = io.BytesIO()
                image.save(img_buffer, format="PNG")
                img_data = img_buffer.getvalue()
                # image_elem = sg.Image(data=img_data)
                
                window["-PDF_IMAGE-"].update(data=img_data)
                window["-TEXT-"].update(text)
            except:
                window["-msg-"].update("Error in processing clipboard")
        else:
            window["-msg-"].update("There is no image in clipboard")
            window["-PDF_IMAGE-"].update()

def process_img_file(img_file,lang):
    image = Image.open(img_file)
    # Convert the image to grayscale
    image = image.convert('L')
    # Use OCR to extract text from the image
    text = pytesseract.image_to_string(image, lang=lang)
    
    img_disp(img_file)
    window["-TEXT-"].update(text)
    
def get_file_name(file_path):
    # this will get the file name of pdf and create same name with docs extension
    file_path_components = file_path.split('/')
    file_name_and_extension = file_path_components[-1].rsplit('.', 1)
    return file_name_and_extension[0]+'.docx'

def visible_fileds(val1,val2,val3):
    window["-msg-"].update("")
    window['-input_m-'].update(val3)
    window['-input_txt-'].update(visible=val1)
    window['-input_file-'].update(visible=val1)
    window['-browse-file-btn-'].update(visible=val1)
    window['-browse-file-btn-'].update(visible=val1)
    window['-res-path-txt-'].update(visible=val2)
    window['-output_dir-'].update(visible=val2)
    window['-browse-fld-btn-'].update(visible=val2)
    window['-str-pg-txt-'].update(visible=val2)
    window['-strt_page-'].update(visible=val2)
    window['-end-pg-txt-'].update(visible=val2)
    window['-end1_page-'].update(visible=val2)
   
# ================================================================================
theme1=["BrightColors","GrayGrayGray","Light Brown 3","LightGreen4","LightGrey4","DarkBlack"]
sg.theme(theme1[2])
# Left & right column in window
input_col = [[sg.Text('Select PDF/Image file to convert to Text',key='-input_txt-',visible=False)],
            [sg.Text(' ', size=(11, 1),key='-input_m-'), 
             sg.Input(key='-input_file-',visible=False), 
             sg.FileBrowse(key='-browse-file-btn-',visible=False)],
            [sg.Text('Output Folder', size=(11, 1),visible=False, key='-res-path-txt-'), 
             sg.Input(key='-output_dir-',visible=False), 
             sg.FolderBrowse(key='-browse-fld-btn-',visible=False)],
            [sg.Text('Start Page', size=(11, 1),visible=False,key='-str-pg-txt-'), 
             sg.Input(size=(7, 1), key='-strt_page-',visible=False),
             sg.Text('End Page', size=(7, 1),visible=False,key='-end-pg-txt-'), 
             sg.InputText(size=(7, 1), key='-end1_page-',visible=False)]
            ]
button_col = [[sg.Text(' '*5),sg.Submit(button_color="green",size=(10,1)),
              sg.Cancel(button_color="red",size=(10,1))] ]

image_col = [[sg.Text('PDF Page / Image:')],
              [sg.Image(key='-PDF_IMAGE-')]]

txt_col = [[sg.Text('Text from PDF Page/Image')], 
            [sg.Multiline('', size=(50, 35), key="-TEXT-")]]
status = "start"
layout = [ [sg.Text("Select Language:"),
            sg.Radio(" हिंदी   ", "lang",default=True, key="-HIN-"), 
            sg.Radio("हिं+Eng  ", "lang",  key="-HINENG-"),    
            sg.Radio("English ", "lang", key="-ENG-")
           ],
           [sg.Text("Select Input:       "),
            sg.Radio("PDF File    ", "input_type", key="-pdf-", enable_events=True), 
            sg.Radio("Image File  ", "input_type",  key="-img-file-", enable_events=True),    
            sg.Radio("From ClipBoard ", "input_type", default=True,key="-clipbrd-", enable_events=True)
           ],
           [sg.Column(input_col, element_justification='l')],
           [sg.Column(button_col, element_justification='c'), sg.VSeperator(),
             sg.Text(' ',key="-msg-",text_color="red")],           
           [sg.Text("Processing . . . ", key="-progress_txt-",visible=False),
            sg.ProgressBar(10, orientation='h', size=(20, 20), key='-progressbar-',visible=False)],
           
           [sg.Text("_" * 80)],
           [sg.Column(image_col, element_justification='c'), sg.VSeperator(),
            sg.Column(txt_col, element_justification='c')]
         ]
# ---------------------------------------------------------------------------------

window = sg.Window('PDF/Image to Text Converter', layout, icon="favicon.ico", size=(800,640))
while True:
    try:
        event, values = window.read()
        if (event == sg.WINDOW_CLOSED) or (event == "Cancel"):
            window.close()
            break
        
        if event == "-pdf-":
            visible_fileds(True, True, " Pdf File")
            
        if event == "-img-file-":
            visible_fileds(True, False, " Image File")
        
        if event == "-clipbrd-":
            visible_fileds(False, False, "  ")
  
        if event == "Submit":
            window["-msg-"].update("")
            lang = get_language(values) 
                    
            if values["-clipbrd-"]:    
                process_clipbrd(lang)
                
            elif values["-img-file-"]:
                img_file = values["-input_file-"]
                # check for png/JPG
                img_type = [".png",".jpg","jpeg"]
                if img_file[-4:].lower() in img_type:
                    process_img_file(img_file,lang)
                else:
                    window["-msg-"].update("File is not PNG/JPG")
                
            else:
                # Get the pdf file 
                pdf_file = values["-input_file-"]
                output_dir = values["-output_dir-"]
                
                if pdf_file[-4:].lower() != ".pdf":
                    window["-msg-"].update("File is not PDF")
                    continue
                
                if values["-strt_page-"].isnumeric():
                    str_page = int(values["-strt_page-"])
                else:
                    window["-msg-"].update("Start Page value is not numeric")
                    continue

                if values["-end1_page-"].isnumeric():
                    end_page = int(values["-end1_page-"])
                else:
                    end_page = str_page
                    
                # Create a new Word document and add the extracted text
                doc = docx.Document()
                # call extract img function
                img_list = extract_img(pdf_file,output_dir, str_page, end_page)
                print(img_list)
                # below code is just for progress bar
                window['-progress_txt-'].update(visible=True)
                window['-progressbar-'].update(visible=True)
                ratio = 10/ (end_page - str_page + 1)
                i = 0   # i is for progress bar count
                for img in img_list: 
                    i = i + 1
                    j = int((i + 1) * ratio)
                    window['-progressbar-'].update_bar(j)
                    # Above code is just for progress bar
                    image = Image.open(img)
                    # Extract text from the image based on the selected language
                    text = extract_text(image, lang)
                    # Add text to Doc file 
                    doc.add_paragraph(text)
                    
                # Save the Word document
                if output_dir > '':
                    doc_file = get_file_name(pdf_file)
                    doc_file = output_dir + "\\" + doc_file
                else:
                    doc_file = pdf_file[:-3] + 'docx'
                doc.save(doc_file)
                
                # Update the image and text on the form
                try:
                    image_elem = sg.Image(data=image)
                    img_disp(img_list[-1])
                    window["-TEXT-"].update(text)
                except:
                    print("Error in updating image and text")
                    continue
    
    except:
        import sys
        exc_type, exc_value, exc_traceback = sys.exc_info()
        # print(f"Error ID: {exc_type.__name__}")
        window["-TEXT-"].update(f"Error Value: {exc_value}")
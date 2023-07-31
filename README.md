# PDFImage_to_Text
PDF/Image to Text Converter using OCR is a Python tool with a PySimpleGUI that extracts text from PDFs, images and clipboard, supporting language selection for accurate recognition. 

##PDF/Image to Text Converter using OCR
This Python program provides a GUI-based tool for converting text from PDF documents and image files using Optical Character Recognition (OCR). The tool is developed using the PySimpleGUI library, making it user-friendly and accessible to users without coding knowledge.

###Features:
Convert PDF Pages to Text: Extract text content from each page of a PDF document and convert it into machine-readable text.
Image to Text Conversion: Import image files and recognize text content from them using OCR.
Language Selection: Choose the language for OCR recognition such as English, Hindi and Eng+Hin, supporting multiple languages for accurate text extraction.
Output to Word Document: The extracted text can be saved in a Word document for further use.

###Installation:
Install Python (https://www.python.org/) and add it to your system's PATH.
Install the required Python libraries:
PySimpleGUI: pip install PySimpleGUI
fitz: pip install PyMuPDF
pytesseract: pip install pytesseract
Pillow: pip install pillow
python-docx: pip install python-docx
Install Tesseract OCR by downloading and installing the appropriate version for your operating system from https://github.com/tesseract-ocr/tesseract.

###Usage:
Run the Python script to launch the GUI window.
Select the desired language for OCR recognition (Hindi, English, or Hindi+English).
Choose the input type (PDF file, image file, or clipboard image).
Depending on the input type, provide the necessary file path, or paste the image from the clipboard.
For PDF files, optionally specify the start and end pages for conversion (default is the entire document).
Click the "Submit" button to process the input and convert text to machine-readable format.
The extracted text will be displayed in the GUI's output window.
If processing a PDF file, the extracted text will be saved to a Word document in the same folder as the input PDF.

###Interface
![image](https://github.com/SHRISTIGUPT/PDFImage_to_Text/assets/91000887/fe01f81e-d672-4cf2-913d-6e8f515c2803)
![image](https://github.com/SHRISTIGUPT/PDFImage_to_Text/assets/91000887/24de6458-cc28-4702-a1e6-6650f9d83410)
![image](https://github.com/SHRISTIGUPT/PDFImage_to_Text/assets/91000887/d36fd68c-1a0b-4cbf-98e6-eb34a9aacdb9)
![image](https://github.com/SHRISTIGUPT/PDFImage_to_Text/assets/91000887/3efce529-b265-4e97-9809-b2b10d5a6250)
![image](https://github.com/SHRISTIGUPT/PDFImage_to_Text/assets/91000887/6216d337-f005-425c-ba14-c2a1a50d7958)

# AgropontosRegex

Agropontos Regex is a small Python program that **extracts geolocation coordinates from PDF files**, eg.: rural property registration documents.

It works for two types of coordinates, **UTM** and **Lat-Long**. And generates a CSV file that can be imported directly to **GIS software**, like QGIS.

The program interface can be used like a notepad to correct any errors or wrong characters brought by the OCR scanning. It also generates a new PDF file correcting the page tilt and rotation.

You need to install the following for Windows:

I recommend using the [Chocolatey](https://chocolatey.org/) package manager to install some of the following: (***Run in an Administrator command prompt***)

- Python 3.8 (64-bit) or later
  - `choco install python3`
- Tesseract 4.1.1 (64-bit) or later
  - `choco install --pre tesseract`
  - You'll also need the [trained data files for Tesseract](https://tesseract-ocr.github.io/tessdoc/Data-Files.html), according to your language 
- Ghostscript 9.50 (64-bit) or later
  - `choco install ghostscript`
- OCRmyPDF 14.2.0 (64-bit) or later
  - `pip install ocrmypdf`
- pypdf 3.9.0 (64-bit) or later
  - `pip install pypdf`

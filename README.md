# Smart OCR Tool

A Python-based desktop application that extracts text from images, PDFs, DOCX, and TXT files using advanced preprocessing techniques.

## Features

* Supports images, PDF, DOCX, and TXT files
* OCR with image preprocessing for better accuracy
* Built-in Text-to-Speech
* Stop & Resume reading (sentence-level precision)
* Copy extracted text
* Save text to file

## Installation

### 1. Install Tesseract OCR

* Windows: https://github.com/UB-Mannheim/tesseract/wiki
* Mac: `brew install tesseract`
* Linux: `sudo apt install tesseract-ocr`

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the application

```bash
python smart_ocr_tool.py
```

## Important

Tesseract OCR must be installed separately for the application to work.

## License

This project is licensed under the MIT License.

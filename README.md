# PDF Transform Script

## Overview

This Python script performs two main tasks:

1. Converts a Microsoft Word `.docx` file to a PDF.
2. Transforms the resulting PDF (or any existing PDF) by **adding an A4-sized blank page next to each page** of the original document. This creates space for handwriting and taking notes.

The script uses **PyPDF2** for PDF manipulation and **win32com.client** for Word-to-PDF conversion.

---

## Features

- **Word to PDF Conversion**: Converts `.docx` files to `.pdf` using Microsoft Word.
- **PDF Transformation**: Adds a blank A4-sized page next to each page of the original PDF for handwriting or notes.
- **Dynamic Output Naming**: The output PDF file is named `my_{input_name}.pdf`, where `{input_name}` is the name of the input Word file without its extension.
- **Command-Line Usage**: Accepts the input Word file as a command-line argument.

---

## Requirements

### Python Dependencies
- `PyPDF2`
- `reportlab`
- `pywin32`

Install them using pip:

```bash
pip install PyPDF2 reportlab pywin32
```

### System Requirements
- **Windows**: The script relies on the `win32com.client` library, which requires Microsoft Word to be installed.
- **Python 3.6 or newer**.

---

## How to Use

### Command-Line Usage
1. Place the script in the same directory as your Word document (or provide the full path to the input file).
2. Run the script with the input file as an argument:

```bash
python script.py <input_file.docx>
```

### Example
```bash
python script.py CS4225_tutorial1.docx
```

### Result
- The script will:
  - Convert `CS4225_tutorial1.docx` to a PDF.
  - Add a blank A4 page next to each page of the PDF.
  - Save the transformed PDF as `my_CS4225_tutorial1.pdf` in the same directory.

---

## File Structure
### Input:
- **Word Document**: A `.docx` file to be converted and transformed.

### Output:
- **Transformed PDF**: A PDF file with blank A4 pages added for notes, saved as `my_{input_name}.pdf`.

---

## Notes
- Ensure the input file is a valid `.docx` file.
- Temporary files (e.g., intermediate PDF files) are cleaned up after processing.
- The output PDF will have the original content on the left and blank A4-sized pages on the right for notes.

---

## Troubleshooting

### Common Errors
1. **"Input file must be a .docx Word document."**
   - Ensure the input file is a `.docx` file.

2. **"Word application error during conversion."**
   - Check if Microsoft Word is installed and properly configured on your system.

3. **"Permission denied"**
   - Ensure you have write permissions for the directory where the output PDF will be saved.

### Debugging
- Enable visibility of the Word application for debugging by modifying the script:
  ```python
  word.Visible = True
  ```

---

## License
This script is distributed under the MIT License. Feel free to use, modify, and distribute it as needed.

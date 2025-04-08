# PPTX to PDF Converter using Microsoft PowerPoint COM (Windows Only)

This Python script batch-converts all `.pptx` files in a given input directory to `.pdf` files using **Microsoft PowerPoint's COM Automation**.  
The converted PDF files are saved in the specified output directory with the same base filename.

> ⚠ This script only works on **Windows** and requires **Microsoft PowerPoint** to be installed.

---

## ✅ Requirements

- Windows OS
- Microsoft PowerPoint (installed and licensed)
- Python 3.x (CPython)
- `comtypes` library

---

## 🚀 How to Use (Step-by-step)

### 1. Clone this repository in Visual Studio Code

Open terminal or use Git panel in VSCode:

```bash
git clone https://github.com/jaehunshin-git/pptxtopdf.git
cd pptx-to-pdf-converter
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

If `requirements.txt` is not available, manually install:

```bash
pip install comtypes
```

### 3. Set your folder paths

Open the Python file (e.g., `convert.py`) and **modify these lines** at the bottom:

```python
input_folder = r"PASTE_YOUR_INPUT_FOLDER_PATH_HERE"
output_folder = r"PASTE_YOUR_OUTPUT_FOLDER_PATH_HERE"
```

Example:

```python
input_folder = r"C:\Users\YourName\Documents\pptx_folder"
output_folder = r"C:\Users\YourName\Documents\pdf_output"
```

### 4. Run the script

Run the script in the terminal:

```bash
python pptx_to_pdf.py
```

You should see conversion logs in the terminal, and `.pdf` files will appear in the output folder.

---

## 📂 Example Directory Structure

```
pptx-to-pdf/
 ├─ pptx_to_pdf.py
 ├─ requirements.txt
 ├─ README.md
 ├─ input_dir/
 │   ├─ sample1.pptx
 │   ├─ sample2.pptx
 ├─ output_dir/
     ├─ sample1.pdf
     ├─ sample2.pdf
```

---

## 📄 Notes

- This script uses PowerPoint itself in the background, so fidelity of the conversion is very high.
- All converted PDFs will have the **same filename** as the original, just with `.pdf` extension.
- Existing files with the same name in the output directory will be **overwritten** without warning.
- Make sure PowerPoint is **not already open** to avoid conflicts.

---

## 📜 License

This project is licensed under the MIT License.

This project uses the [comtypes](https://github.com/enthought/comtypes) library, which is licensed under the MIT License.

Note: This script automates Microsoft PowerPoint through COM automation.  
**Microsoft PowerPoint must be installed and properly licensed** on the system where this script is executed.  
Microsoft and PowerPoint are trademarks of the Microsoft group of companies.


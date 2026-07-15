# 📄 Excel to PDF Converter & Formatter (Special Edition)

A powerful, multi-threaded Python GUI application designed to parse, clean, and format industrial log sheets from Microsoft Excel (`.xlsx`, `.xls`) into structured, publication-quality PDF reports. 

This application is customized for heat treatment monitoring, specifically tracking sensor temperatures and highlighting critical process events.

---

## 🚀 Key Features

* **Dynamic Column Sizing (Anti-Overlap)**:
  Analyzes cell text lengths dynamically to scale column widths proportionally within A4 boundaries. This completely prevents text truncation or horizontal overlap.
* **Smart Orientation Switching**:
  Detects column density and automatically switches the PDF page layout between **Portrait** and **Landscape** (for datasets with > 8 columns).
* **Precise Date & Time Normalization**:
  Auto-detects Date and Time columns by header name scans. Converts date fields into clean `d.m.Y` format (e.g., `27.3.2025`) and time fields into `H:M:S.f` with single decimal place precision (e.g., `00:15:54.9`).
* **Heat Treatment Target Event Highlighting**:
  Examines all temperature sensor/probe columns (excluding metadata). For each of the 16 probes, it identifies the exact cell where the probe reaches `56°C` for the first time and highlights **ONLY** those specific cells in yellow.
* **Data Integrity Preserved**:
  The application **does not modify the original data** in any way. It is purely designed for verification and marking purposes to control field operations and convert the results into a PDF report.
* **Fast Concurrent Batch Processing**:
  Utilizes `ThreadPoolExecutor` to convert dozens of log sheets concurrently while displaying a real-time progress bar.
* **Modern GUI & Packaging**:
  Features a clean interface with a scaled company logo, high-resolution multi-size application icons (up to `256x256` pixels), and a detailed error diagnostics dialog.

---

## ⚙️ Technologies

* **Python 3.12**
* **Pandas** & **OpenPyXL** – Excel sheet data extraction and cleaning.
* **FPDF2** – Layout creation, custom Unicode TTF font embedding, and cell rendering.
* **Tkinter** – Desktop Graphical User Interface.
* **Pillow** – High-resolution image scaling and processing.
* **PyInstaller** – Standalone executable compilation.

---

## 📥 Installation

### 1. Prerequisites
Ensure you have Python 3.10+ installed.

### 2. Install Dependencies
Install all required libraries using pip:
```sh
pip install pandas fpdf2 pillow python-dateutil openpyxl
```

### 3. Running the App locally
Run the script directly:
```sh
python src/ExcelToPdfSpecial.py
```

---

## 🖥️ Usage

1. **Select Excel Files**: Click **"Dosya Seç"** and select one or multiple `.xlsx` or `.xls` log sheets.
2. **Stop at Blank (Optional)**: Check **"Boş satırdan sonra dur"** if you want the parser to truncate the log when a sensor column registers an empty cell.
3. **Convert**: Click **"Dönüştür"**, choose the destination folder, and the application will convert the files in the background.
4. **View Outputs**: The formatted PDF documents with highlighted process rows will be generated in your selected directory.

---

## 🛠️ Standalone Compilation (EXE)

To bundle the application into a single, standalone Windows executable (`.exe`) with embedded high-resolution icons and fonts, execute PyInstaller:
```powershell
.\venv\Scripts\pyinstaller --noconfirm --onefile --windowed --icon="assets/favicon.ico" --add-data "assets;assets" src/ExcelToPdfSpecial.py
```
The output executable will be generated inside the `dist/` directory.

---

## 📩 Contact & Repository

* **Developer**: [Murat Can](https://github.com/mrt-cn)
* **GitHub Repository**: [ExcelToPdfSpecial](https://github.com/mrt-cn/ExcelToPdfSpecial)

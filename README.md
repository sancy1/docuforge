<div align="center">

<img width="154" height="117" alt="image" src="https://github.com/user-attachments/assets/8efa49e9-d4f7-43d4-8aee-35717792b6e3" />

# DocuForge Converter Pro

### Professional Document Conversion Suite

</div>

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Desktop App](https://img.shields.io/badge/Desktop-Tkinter-green)
![License](https://img.shields.io/badge/License-MIT-lightgrey)
![Status](https://img.shields.io/badge/Status-Active-success)

**Professional Document Conversion Suite — Local-First, Zero Cloud, Maximum Fidelity**

DocuForge Converter Pro is a powerful desktop application designed to streamline document conversion while preserving layout, formatting, fonts, tables, and overall document integrity.

Built with Python and a premium dark-themed interface powered by **ttkbootstrap**, DocuForge bridges the gap between document formats such as Word, PDF, Excel, CSV, TXT, ODT, and RTF without sacrificing quality.

Unlike many online converters, **DocuForge operates entirely on your machine**. No files are uploaded, no cloud processing occurs, and no sensitive information ever leaves your computer.

---

# 📸 Screenshots

> Insert screenshots directly below. Keeping screenshots near the top ensures GitHub renders them correctly and gives visitors an immediate visual understanding of the application.

![Main Dashboard](INSERT_SCREENSHOT_URL_HERE)

![Batch Conversion](INSERT_SCREENSHOT_URL_HERE)

![Dark Theme Interface](INSERT_SCREENSHOT_URL_HERE)

---

# 🚀 Key Features

## Universal Document Conversion

Convert seamlessly between:

| Source Format | Supported Outputs             |
| ------------- | ----------------------------- |
| DOCX          | PDF, TXT, XLSX, CSV, ODT, RTF |
| PDF           | DOCX, TXT, XLSX, CSV          |
| XLSX          | PDF, DOCX, CSV, TXT           |
| CSV           | PDF, DOCX, XLSX, TXT          |
| TXT           | DOCX, PDF, XLSX, CSV          |
| ODT           | PDF                           |
| RTF           | PDF                           |

---

## 🎯 Fidelity-First Architecture

DocuForge prioritizes document accuracy.

### Features

* Font Preservation
* Layout Preservation
* Embedded Font Support
* Table Structure Retention
* Image Preservation
* Margin Preservation
* Header & Footer Preservation

### LibreOffice Integration

When available, DocuForge uses LibreOffice Headless Mode for pixel-perfect document rendering.

Benefits include:

* DOCX → PDF fidelity comparable to Microsoft Word
* ODT → PDF conversion support
* Better font handling
* Improved image rendering

---

## 📊 Smart Data Processing

### Spreadsheet Preservation

When converting spreadsheet-based formats:

* Auto-sized columns
* Clean table rendering
* Professional PDF table layouts
* Landscape support for wide datasets
* Header styling
* Alternating row colors

### Data Sanitization

Automatically removes:

* NaN values
* Unnamed columns
* Empty rows
* Empty columns
* Corrupted metadata

### Encoding Detection

Supported encodings:

* UTF-8
* UTF-8-SIG
* Latin-1
* CP1252

---

## 📁 Folder Structure Preservation

Batch conversion supports:

* Single files
* Multiple files
* Entire folders
* Nested directories

Features:

* Original hierarchy retained
* Automatic output grouping
* Consistent file naming
* Bulk conversion workflows

---

## 🎨 Premium Desktop Experience

Built with a high-contrast dark theme using ttkbootstrap.

Features include:

* Modern responsive interface
* Conversion queue management
* Progress indicators
* Detailed logging
* Status tracking
* Error reporting
* Success notifications

---

# 🛠 Technology Stack

## Core Framework

| Technology   | Purpose             |
| ------------ | ------------------- |
| Python 3.10+ | Application Runtime |
| Tkinter      | Desktop GUI         |
| ttkbootstrap | Modern UI Styling   |

---

## Document Processing

| Library     | Purpose                |
| ----------- | ---------------------- |
| python-docx | DOCX Processing        |
| docx2txt    | Text Extraction        |
| PyMuPDF     | PDF Processing         |
| reportlab   | PDF Generation         |
| pdfkit      | HTML to PDF Conversion |

---

## Data Processing

| Library  | Purpose           |
| -------- | ----------------- |
| pandas   | Data Manipulation |
| openpyxl | Excel Processing  |
| csv      | CSV Processing    |

---

## System Integration

| Tool        | Purpose                        |
| ----------- | ------------------------------ |
| LibreOffice | High-Fidelity Rendering        |
| subprocess  | Headless Automation            |
| threading   | Background Tasks               |
| tempfile    | Temporary Workspace Management |

---

# ⚙ Installation

## Prerequisites

* Python 3.10+
* pip
* LibreOffice 7.0+ (Recommended)

---

## Clone Repository

```bash
git clone https://github.com/yourusername/docuforge.git

cd docuforge
```

---

## Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Run Application

```bash
python docuforge.py
```

---

# 📖 Usage Guide

## Quick Start

### Step 1

Launch the application:

```bash
python docuforge.py
```

### Step 2

Add files or folders using:

* 📁 Add Files
* 📂 Add Folder

### Step 3

Choose the desired output format.

### Step 4

Click:

```text
⚡ Start Conversion
```

### Step 5

Select an output location and monitor progress.

---

# 🔤 Font Preservation

## The Problem

Traditional document converters frequently replace missing fonts, resulting in formatting inconsistencies and visual degradation.

---

## The Solution

DocuForge forces PDF font embedding whenever possible.

Benefits include:

* Consistent rendering
* Improved portability
* Professional output
* Cross-platform compatibility

---

# 🔄 Conversion Matrix

| From → To | PDF | DOCX | XLSX | CSV | TXT | ODT | RTF |
| --------- | --- | ---- | ---- | --- | --- | --- | --- |
| DOCX      | ✅   | ✅    | ✅    | ✅   | ✅   | ✅   | ✅   |
| PDF       | ✅   | ✅    | ✅    | ✅   | ✅   | ❌   | ❌   |
| XLSX      | ✅   | ✅    | ✅    | ✅   | ✅   | ❌   | ❌   |
| CSV       | ✅   | ✅    | ✅    | ✅   | ✅   | ❌   | ❌   |
| TXT       | ✅   | ✅    | ✅    | ✅   | ✅   | ❌   | ❌   |
| ODT       | ✅   | ❌    | ❌    | ❌   | ❌   | ✅   | ❌   |
| RTF       | ✅   | ❌    | ❌    | ❌   | ❌   | ❌   | ✅   |

---

# ❓ FAQ

### Does DocuForge require LibreOffice?

No. However, LibreOffice significantly improves document fidelity during conversion.

### Is cloud processing used?

No. DocuForge is completely local-first.

### Can I batch-convert files?

Yes. Entire directories can be processed while preserving folder structure.

### Is there a file size limit?

No hard limit exists, though larger files may require additional processing time.

---

# 🗺 Roadmap

Planned features include:

* OCR for scanned PDFs
* Cloud Storage Integrations
* Document Merging
* Document Splitting
* Image Extraction
* Batch Renaming
* CLI Automation
* Template-Based Exports

---

# 👨‍💻 Author

## Alexander Cyril

Professional Software Developer

* Desktop Applications
* Document Processing Systems
* Cloud & Backend Engineering
* Full-Stack Development

📧 [alexander.s.cyril@gmail.com](mailto:alexander.s.cyril@gmail.com)

---

# 📄 License

This project is licensed under the MIT License.

See the LICENSE file for full details.

---

**DocuForge Converter Pro** — Transform Documents with Confidence ⚡

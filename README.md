# COMP-4710-AIBCR

## Overview
For our COMP 4710 class project, we are building a toolset to automate the tedious process of reviewing government and commercial contracts for our sponsor. 

Currently, reviewing these documents means manually reading hundreds of pages of legalese, searching for alphanumeric federal codes, and cross-referencing them against massive Excel databases to see if they are acceptable. It takes hours. This project automates that entire workflow, saving our sponsor massive amounts of time while reducing human error.

## Phase 1: The Legal Code Flagger
The first Python script (`document_flagger.py`) handles standard federal solicitations. It acts as an incredibly advanced text parser and cross-referencer.

**What it does:**
* **Multi-Format Reading:** It reads native digital PDFs, modern Word documents (`.docx`), and even uses an OCR fallback (Tesseract) to read scanned image PDFs.
* **Database Merging:** It automatically loads and stitches together multiple agency databases (FAR, DFARS, NASA, etc.) from our `Database` folder into one master dictionary.
* **Color-Coded Output:** It generates an Excel Compliance Matrix and automatically color-codes the rows based on the sponsor's legal rubric (Green = "OK", Yellow = "C" / Conditional, Red = "Remove").
* **Auto-Highlighting:** It physically draws bright yellow highlight boxes over every found clause directly inside a new "Executed" copy of the PDF or Word document so the human reviewer can easily spot them.

---

## Phase 2: AI Semantic Contract Reviewer
While Phase 1 works perfectly for federal contracts that rely on strict alphanumeric codes, it does not work for custom, paragraph-based commercial contracts or private grants (e.g., Korea Foundation agreements). 

To solve this, we built a second tool: **Script #2 (`ai_reviewer.py`)**. 

Instead of searching for specific code patterns (regex), this script uses Large Language Model (LLM) integration (Google Gemini 2.5 Flash) to perform **complex semantic detection** while strictly adhering to the sponsor's legal boundaries.

**What it does:**
* **Dynamic Legal Playbook:** The script dynamically reads the sponsor's `Contract Ts&Cs Matrix.xlsm` file. If the sponsor ever updates their rules or adds new tabs, the AI instantly learns them without requiring any code changes.
* **Natural Language Reading:** The AI reads the English paragraphs of custom commercial contracts to understand the context of the terms.
* **Cited Redlining & Risk Reporting:** The output is a cleanly formatted Word document containing a summary of risky terms, why they violate Auburn University's rules (e.g., sovereign immunity), suggested legal redlines to send back to the vendor, and **explicit citations** pointing the reviewer exactly to the tab in the matrix where the rule originated.

---

### Folder Structure Note
*Note to Graders/Reviewers: The actual contract files and database `.csv` / `.xlsm` files are ignored via `.gitignore` to protect sensitive client data. If you clone this repo, the `Database`, `Solicitations`, and `Output` folders will be empty.*

### How to Run Locally (Mac)

1. Install system dependencies for the OCR engine:
    brew install tesseract
    brew install poppler

2. Install the required Python libraries:
    python -m pip install pandas openpyxl pdfplumber pytesseract pdf2image python-docx PyMuPDF google-genai

3. Get a free Google Gemini API Key from Google AI Studio and paste it into line 17 of `ai_reviewer.py`.

4. Place your master database files in the `Database/` folder and your contracts in the `Solicitations/` folder.

5. Run the scripts depending on your needs:
    cd Code
   
    # For Federal Contracts:
    python document_flagger.py
   
    # For Commercial/Custom Contracts:
    python ai_reviewer.py

---

### How to Run Locally (Windows)

1. Install System Dependencies (Required for reading scanned PDFs):
    * **Tesseract OCR:** Download the Windows installer from the [UB-Mannheim GitHub page](https://github.com/UB-Mannheim/tesseract/wiki). Run the installer and note the installation path (usually `C:\Program Files\Tesseract-OCR`).
    * **Poppler:** Download the latest Windows binary from the [Poppler for Windows repository](https://github.com/oschwartz10612/poppler-windows/releases). Extract the folder and place it somewhere permanent (like `C:\poppler`).
    * **Crucial Step:** You must add the `bin` folders for BOTH of these programs to your Windows System `PATH` Environment Variable so Python can find them. 

2. Install the required Python libraries:
    Open Command Prompt or PowerShell and run:
    python -m pip install pandas openpyxl pdfplumber pytesseract pdf2image python-docx PyMuPDF google-genai

3. Get a free Google Gemini API Key from Google AI Studio and paste it into line 17 of `ai_reviewer.py`.

4. Place your master database files in the `Database/` folder and your contracts in the `Solicitations/` folder.

5. Run the scripts depending on your needs:
    cd Code
   
    # For Federal Contracts:
    python document_flagger.py
   
    # For Commercial/Custom Contracts:
    python ai_reviewer.py
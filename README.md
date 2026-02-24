# COMP-4710-AIBCR

## Overview
For our COMP 4710 class project, we are building a toolset to automate the tedious process of reviewing government and commercial contracts for our sponsor. 

Currently, reviewing these documents means manually reading hundreds of pages of legalese, searching for alphanumeric federal codes, and cross-referencing them against massive Excel databases to see if they are acceptable. It takes hours. This project automates that entire workflow, saving our sponsor massive amounts of time.

## Phase 1: The Legal Code Flagger (Current State)
The current Python script (`document_flagger.py`) handles standard federal solicitations. It acts as an incredibly advanced text parser and cross-referencer.

**What it does:**
* **Multi-Format Reading:** It reads native digital PDFs, modern Word documents (`.docx`), and even uses an OCR fallback (Tesseract) to read scanned image PDFs.
* **Database Merging:** It automatically loads and stitches together multiple agency databases (FAR, DFARS, NASA, etc.) from our `Database` folder into one master dictionary.
* **Color-Coded Output:** It generates an Excel Compliance Matrix and automatically color-codes the rows based on the sponsor's legal rubric (Green = "OK", Yellow = "C" / Conditional, Red = "Remove").
* **Auto-Highlighting:** It physically draws bright yellow highlight boxes over every found clause directly inside a new "Executed" copy of the PDF or Word document so the human reviewer can easily spot them.

### Folder Structure Note
*Note to Graders/Reviewers: The actual contract files and database `.csv` files are ignored via `.gitignore` to protect sensitive client data. If you clone this repo, the `Database`, `Solicitations`, and `Output` folders will be empty.*

### How to Run Locally (Mac)
1. Install system dependencies for the OCR engine:
   ```bash
   brew install tesseract
   brew install poppler
   ```
2. Install the required Python libraries:
   ```bash
   pip install pandas openpyxl pdfplumber pytesseract pdf2image python-docx PyMuPDF
   ```
3. Place your master database files in the `Database/` folder and your contracts in the `Solicitations/` folder.
4. Run the script:
   ```bash
   cd Code
   python document_flagger.py
   ```

---

## Phase 2: AI Semantic Detection (Future Goals)
While Phase 1 works perfectly for federal contracts that rely on strict alphanumeric codes, it does not work for custom, paragraph-based commercial contracts or private grants (e.g., Korea Foundation agreements). 

To solve this, our next goal is to build **Script #2: The AI Contract Reviewer**.

Instead of searching for specific code patterns (regex), this script will use Large Language Model (LLM) integration (like the Google Gemini API) to perform **complex semantic detection**. 

**Future Features:**
* **Natural Language Reading:** The AI will read the English paragraphs of a custom contract to understand the context of the terms.
* **Custom Legal Playbooks:** We will program the AI with our sponsor's specific non-negotiable rules. For example: automatically flagging any clause that attempts to subject Auburn University to a foreign jurisdiction, ensuring the University's sovereign immunity as an instrumentality of the State of Alabama is protected.
* **Automated Redlining:** The output will be a Word document containing a summary of risky terms, why they violate the playbook, and suggested legal redlines to send back to the vendor.
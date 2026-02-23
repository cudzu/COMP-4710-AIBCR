"""
=============================================================================
Project: Automated Compliance Matrix Generator
Description: This script reads government contract PDFs, finds the federal 
             clauses, and builds a spreadsheet matching those clauses to our 
             master database.
=============================================================================

=============================================================================
Generative AI Use: 
This code was commented and structured with the help of Google Gemini 3.1
to make it easier to understand and maintain. The AI helped explain the 
purpose of each function and the overall flow of the script in simple terms, 
so that even someone new to programming can follow along.
=============================================================================

"""

import os               # Helps the script find folders and files on the computer
import re               # Helps find specific text patterns (like clause numbers)
import pandas as pd     # Used to read and write Excel and CSV files
import pdfplumber       # Used to read text from normal digital PDFs
import pytesseract      # Used to read text from scanned pictures/documents
from pdf2image import convert_from_path  # Turns PDF pages into pictures if needed
from datetime import datetime            # Adds the current time to our output files

# =============================================================================
# --- Folder Setup ---
# =============================================================================

# The script looks for this exact column name to match everything up.
CLAUSE_COL_NAME = 'Clause' 

# Find exactly where this code is saved on the computer.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Look one folder up to find the main project folder.
if os.path.basename(SCRIPT_DIR).lower() == 'code':
    PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
else:
    PROJECT_DIR = SCRIPT_DIR

# Set the paths for where our data lives and where to save the results.
DATABASE_DIR = os.path.join(PROJECT_DIR, 'Database')
SOLICITATIONS_DIR = os.path.join(PROJECT_DIR, 'Solicitations')
OUTPUT_DIR = os.path.join(PROJECT_DIR, 'Output')

# =============================================================================
# --- Project Functions ---
# =============================================================================

def clean_headers(columns):
    """
    Cleans up messy column names in the Excel files. 
    Sometimes the files have weird spaces or asterisks in the headers. 
    This fixes them so they all match perfectly.
    """
    clean_cols = []
    for col in columns:
        # Remove line breaks, asterisks, and extra spaces
        c = str(col).replace('\n', ' ').replace('*', '').strip()
        c = re.sub(r'\s+', ' ', c)
        clean_cols.append(c)
    return clean_cols

def load_databases(db_dir):
    """
    Reads all our separate agency files (FAR, DFARS, NASA, etc.) from the 
    Database folder and combines them into one big list for the script to use.
    """
    print(f"Loading database files from: {db_dir}...")
    all_dataframes = [] 
    
    if not os.path.exists(db_dir):
        print(f"Error: Database folder '{db_dir}' not found.")
        return None

    # Go through every file in the database folder
    for filename in os.listdir(db_dir):
        filepath = os.path.join(db_dir, filename)
        
        # Skip hidden files or files we don't want to scan
        if filename.startswith('~') or filename.startswith('.') or 'Definitions' in filename or filename == 'Contract Ts&Cs Matrix.xlsm':
            continue 

        try:
            # Open the file depending on if it's Excel or CSV
            if filename.lower().endswith(('.xlsx', '.xlsm')):
                df = pd.read_excel(filepath, engine='openpyxl')
            elif filename.lower().endswith('.xls'):
                df = pd.read_excel(filepath)
            elif filename.lower().endswith('.csv'):
                try:
                    df = pd.read_csv(filepath, encoding='utf-8', on_bad_lines='skip')
                except UnicodeDecodeError:
                    df = pd.read_csv(filepath, encoding='latin1', on_bad_lines='skip')
            else:
                continue 
            
            # Clean up the column names
            df.columns = clean_headers(df.columns)
            
            # Filter out junk rows (like instructions or blank lines)
            if CLAUSE_COL_NAME in df.columns:
                df[CLAUSE_COL_NAME] = df[CLAUSE_COL_NAME].astype(str).str.strip()
                # Make sure the clause actually has a number in it
                df = df[df[CLAUSE_COL_NAME].str.contains(r'\d', na=False)]
                # Make sure the clause isn't an entire paragraph
                df = df[df[CLAUSE_COL_NAME].str.len() < 30]
                
                all_dataframes.append(df)
            
        except Exception as e:
            print(f"  ! Error loading {filename}: {e}")

    if not all_dataframes:
        return None

    # Combine everything together
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    print(f"Successfully built a master database of {len(combined_df)} total clauses.")
    return combined_df

def extract_text_from_pdf(pdf_path):
    """
    Tries to read the PDF normally. If it realizes the PDF is just a scanned 
    picture, it takes extra steps to read the text from the images.
    """
    text = ""
    print(f"Extracting text from: {os.path.basename(pdf_path)}...")
    
    # Try reading the text the fast, normal way
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"  ! Error reading PDF {pdf_path}: {e}")
        return None

    # If we barely found any text, it's probably a scanned document.
    if len(text.strip()) < 50:
        print("  ! No text found. Looks like a scan. Reading images now (this might take a minute)...")
        try:
            # Turn the PDF pages into pictures
            images = convert_from_path(pdf_path)
            ocr_text = ""
            
            # Read the text out of each picture
            for i, image in enumerate(images):
                print(f"    - Reading page {i + 1} of {len(images)}...")
                ocr_text += pytesseract.image_to_string(image) + "\n"
            
            text = ocr_text
            
            if not text.strip():
                print("  ! Warning: Could not find any text in the scan either. Skipping.")
                return None
                
        except Exception as e:
            print(f"  ! Image reading failed: {e}")
            return None

    return text

def find_clauses_from_db(text, master_clauses):
    """
    Looks through the text we pulled from the PDF and checks if any of our 
    database clauses are in it.
    """
    found_clauses = []
    for clause in master_clauses:
        # Check for the exact clause number (so we don't accidentally match part of another number)
        pattern = r'\b' + re.escape(clause) + r'\b'
        if re.search(pattern, text):
            found_clauses.append(clause)
            
    return sorted(found_clauses)

def generate_compliance_matrix(found_clauses, master_df):
    """
    Builds the final Excel spreadsheet by matching the clauses we found in the 
    PDF with their full details from our database.
    """
    if master_df is None:
        return None

    matrix_rows = []
    headers = master_df.columns.tolist()

    print(f"Cross-referencing {len(found_clauses)} clauses with the combined database...")

    # Grab the row information for every clause we found
    for clause in found_clauses:
        matching_rows = master_df[master_df[CLAUSE_COL_NAME] == clause]
        for _, row in matching_rows.iterrows():
            matrix_rows.append(row.tolist())

    # Put it all into a clean spreadsheet format and sort it
    matrix_df = pd.DataFrame(matrix_rows, columns=headers)
    matrix_df = matrix_df.sort_values(by=CLAUSE_COL_NAME).reset_index(drop=True)
    return matrix_df

# =============================================================================
# --- Main Script ---
# =============================================================================

def main():
    print("\n--- Starting Automated Compliance Matrix Generator ---")

    # Make sure our folders exist
    os.makedirs(DATABASE_DIR, exist_ok=True)
    os.makedirs(SOLICITATIONS_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Load all the database files
    master_df = load_databases(DATABASE_DIR)
    if master_df is None:
        print("Exiting due to error loading databases.")
        return

    # Make a list of all the known clauses to search for
    known_clauses = master_df[CLAUSE_COL_NAME].unique().tolist()

    # Find all the PDFs we need to scan
    pdf_files = [f for f in os.listdir(SOLICITATIONS_DIR) if f.lower().endswith('.pdf')]

    if not pdf_files:
        print(f"\nNo PDF files found in {SOLICITATIONS_DIR}.")
        return

    # Process each PDF one by one
    for pdf_file in pdf_files:
        print(f"\n--- Processing: {pdf_file} ---")
        pdf_path = os.path.join(SOLICITATIONS_DIR, pdf_file)

        # Get the text from the PDF
        text = extract_text_from_pdf(pdf_path)
        if not text:
            continue

        # Find the matching clauses
        found_clauses = find_clauses_from_db(text, known_clauses)
        print(f"Found {len(found_clauses)} unique federal clauses.")

        if not found_clauses:
            print("Warning: No matching federal clauses found in text. Skipping file save.")
            continue

        # Create the final spreadsheet
        compliance_matrix_df = generate_compliance_matrix(found_clauses, master_df)

        # Name the file with a timestamp so we don't overwrite older files
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_filename = f"Compliance_Matrix_{os.path.splitext(pdf_file)[0]}_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        
        # Save the file to the Output folder
        try:
            compliance_matrix_df.to_excel(output_path, index=False)
            print(f"Successfully saved compliance matrix to:\n  -> {output_path}")
        except Exception as e:
            print(f"Error saving output file {output_path}: {e}")

    print("\n--- All tasks completed. ---")

# Run the main function when the script starts
if __name__ == '__main__':
    main()
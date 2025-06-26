import os
import PyPDF2
import pandas as pd
from tqdm import tqdm
import re

def extract_year_from_folder(folder_name):
    # Extract year from folder name (ex., "2005 ICRA Articles" → 2005)
    match = re.match(r'^(19\d{2}|20\d{2}|2100)', folder_name)
    return int(match.group(0)) if match else None

def count_words_in_pdf(pdf_path, words_to_count):
    counts = {word.lower(): 0 for word in words_to_count}
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = "".join([page.extract_text() or "" for page in reader.pages]).lower()
            for word in words_to_count:
                counts[word.lower()] = text.count(word.lower())
        return counts, None
    except Exception as e:
        return None, str(e)

def process_pdfs_with_year(root_folder, output_file="pdf_word_counts_with_year.xlsx"):
    words_to_count = [
        "bipedal", "walking", "human-like", "humanoid", "rescue robot", "underwater",
        "pipe inspection", "single-wheel", "slam", "legged", "quadrupedal",
        "autonomous navigation", "self-driving", "uav", "drone", "autonomous vehicle",
        "field robotics", "ground vehicle", "robotic wheelchair", "agv", "mobile robot",
        "aerial robot", "quadruped", "locomotion", "urban environments", "climbing", "jumping"
    ]
    
    data = []
    error_log = []
    pdf_files = []
    
    # Gather PDF files first for accurate progress bar
    for foldername, _, filenames in os.walk(root_folder):
        if "compressed" in foldername.lower():
            continue
        for filename in filenames:
            if filename.lower().endswith('.pdf'):
                pdf_files.append((foldername, filename))
    
    # Process files with progress bar
    for foldername, filename in tqdm(pdf_files, desc="Processing PDFs", unit="file"):
        pdf_path = os.path.join(foldername, filename)
        relative_folder = os.path.relpath(foldername, root_folder)
        counts, error = count_words_in_pdf(pdf_path, words_to_count)
        
        if counts:
            entry = {
                "PDF Name": filename,
                "Folder": relative_folder,
                "Year": extract_year_from_folder(os.path.basename(foldername))
            }
            entry.update(counts)
            data.append(entry)
        else:
            error_log.append(f"{pdf_path} | Error: {error}")
    
    # Save results
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    
    if error_log:
        with open("processing_errors.log", "w") as f:
            f.write("\n".join(error_log))
    
    print(f"\nProcessed {len(data)} PDFs. Results saved to {output_file}")
    if error_log:
        print(f"Encountered {len(error_log)} errors (see processing_errors.log)")
        
PDF_FOLDER = r"C:\Users\mateo\ICRA Articles" # Update as needed

if __name__ == "__main__":
    process_pdfs_with_year(PDF_FOLDER)

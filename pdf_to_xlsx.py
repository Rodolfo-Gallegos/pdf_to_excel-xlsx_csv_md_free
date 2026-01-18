import os
import sys
import argparse
import logging
import time
from typing import List, Optional

import pdfplumber
import pandas as pd
from google import genai
from dotenv import load_dotenv
from PIL import Image

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# Silence SDK internal logging (suppress "AFC is enabled" and other noise)
for logger_name in ["google", "google.genai", "urllib3"]:
    logging.getLogger(logger_name).setLevel(logging.WARNING)

# Load API Key
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(SCRIPT_DIR, "api_key.env")
load_dotenv(ENV_PATH)
api_key = os.getenv("API_KEY")

if not api_key:
    logger.error("\n\tAPI_KEY not found in api_key.env")
    sys.exit(1)

try:
    client = genai.Client(api_key=api_key)
except Exception as e:
    logger.error(f"\n\tFailed to initialize Gemini client: {e}")
    sys.exit(1)

def get_gemini_response(prompt: str, content: Optional[Image.Image] = None) -> str:
    """
    Generic helper to call Gemini with retry logic for quota limits.
    
    Args:
        prompt: The text prompt for the model.
        content: Optional image data.
        
    Returns:
        The text response from the model.
    """
    gemini_model = 'gemini-3-flash-preview'
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if content:
                response = client.models.generate_content(
                    model=gemini_model,
                    contents=[prompt, content]
                )
            else:
                response = client.models.generate_content(
                    model=gemini_model,
                    contents=prompt
                )
            
            if response and hasattr(response, 'text') and response.text:
                return response.text
            return ""
        except Exception as e:
            if "429" in str(e) and attempt < max_retries - 1:
                wait_time = (attempt + 1) * 10
                logger.warning(f"Quota limit reached, waiting {wait_time}s... (Attempt {attempt+1}/{max_retries})")
                time.sleep(wait_time)
            else:
                logger.error(f"Error calling Gemini API: {e}")
                raise e
    return ""

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans and normalizes DataFrame content for programmatic use.
    - Removes currency symbols ($, €, etc.)
    - Removes commas from numbers
    - Strips whitespace
    - Standardizes empty strings
    
    Args:
        df: Input DataFrame.
        
    Returns:
        A cleaned version of the DataFrame.
    """
    def clean_cell(val):
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        # Remove currency symbols and formatting commas
        s = s.replace('$', '').replace('€', '').replace(',', '')
        # Handle cases where leading zeros are literal (IDs) but try to keep numbers as numbers
        # If it looks like a number, return it as such
        try:
            if '.' in s:
                return float(s)
            return int(s)
        except ValueError:
            return s

    return df.apply(lambda col: col.map(clean_cell))

def parse_markdown_to_df(md_text: str) -> pd.DataFrame:
    """
    Converts Markdown tables to pandas DataFrames.
    
    Args:
        md_text: String containing potential Markdown tables.
        
    Returns:
        A pandas DataFrame containing the extracted data.
    """
    if not md_text or not isinstance(md_text, str):
        return pd.DataFrame()
        
    lines = md_text.strip().split('\n')
    data = []
    for line in lines:
        if '|' in line:
            cells = [c.strip() for c in line.split('|')]
            # Remove leading/trailing empty cells from pipe alignment
            if cells and cells[0] == '': cells = cells[1:]
            if cells and cells[-1] == '': cells = cells[:-1]
            # Skip separator rows like |---|---|
            if all(set(c) <= {'-', ':', ' '} for c in cells):
                continue
            if cells:
                data.append(cells)
        else:
            if line.strip():
                data.append([line.strip()])
    
    if not data:
        return pd.DataFrame()
        
    # Standardize column count
    max_cols = max(len(row) for row in data)
    normalized_data = [row + [''] * (max_cols - len(row)) for row in data]
    
    return pd.DataFrame(normalized_data)

def process_page_images(page: pdfplumber.page.Page) -> List[dict]:
    """
    Analyzes a PDF page as an image to extract visual tables and their raw markdown.
    
    Args:
        page: A page object from pdfplumber.
        
    Returns:
        A list of dictionaries with 'df' and 'md' keys.
    """
    logger.info(f"\n\tAnalyzing page {page.page_number} visually...")
    try:
        img = page.to_image(resolution=300).original
    except Exception as e:
        logger.error(f"\n\tFailed to render page {page.page_number}: {e}")
        return []
    
    prompt = """
    Analyze this page and extract ALL tables you see.
    Even if the table looks like a screenshot or an embedded image, extract it.
    Return results strictly in Markdown format.
    Do not include any introductory text, titles outside the table, or comments.
    If no tables are found, return an empty string.
    """
    
    try:
        md = get_gemini_response(prompt, img)
    except Exception:
        return []
    
    # Clean possible markdown code blocks
    md_clean = md.replace("```markdown", "").replace("```", "").strip()
    
    if md_clean:
        df = parse_markdown_to_df(md_clean)
        if not df.empty:
            return [{"df": df, "md": md_clean}]
            
    return []

def main(pdf_files: List[str], output_path: str, save_md: bool = False, save_csv: bool = False, clean: bool = False) -> None:
    """
    Main function to extract tables from PDFs and save to multiple formats.
    
    Args:
        pdf_files: List of paths to PDF files.
        output_path: Path where the resulting Excel file will be saved.
        save_md: Whether to save raw markdown.
        save_csv: Whether to save each table as CSV.
        clean: Whether to normalize/clean the extracted data.
    """
    valid_files = [f for f in pdf_files if os.path.exists(f)]
    if not valid_files:
        logger.error("\n\tNo valid PDF files to process.")
        return

    try:
        excel_name = output_path if output_path.endswith('.xlsx') else f"{output_path}.xlsx"
        writer = pd.ExcelWriter(excel_name, engine='openpyxl')
        
        with writer:
            pd.DataFrame([["\n\tTables extracted from PDF images"]]).to_excel(writer, sheet_name="Summary", index=False, header=False)
            
            for pdf_path in valid_files:
                file_name = os.path.basename(pdf_path)
                logger.info(f"\n\tProcessing {file_name}...")
                
                all_results = []
                with pdfplumber.open(pdf_path) as pdf:
                    for page in pdf.pages:
                        page_results = process_page_images(page)
                        all_results.extend(page_results)
                
                if all_results:
                    # Process DataFrames (Apply cleaning if requested)
                    processed_results = []
                    for res in all_results:
                        df = res['df']
                        if clean:
                            df = normalize_dataframe(df)
                        processed_results.append({"df": df, "md": res['md']})

                    # Combine for Excel
                    all_dfs = [res['df'] for res in processed_results]
                    combined_df = pd.concat(all_dfs, ignore_index=True)
                    sheet_name = os.path.splitext(file_name)[0][:31]
                    combined_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    
                    # Save MD if requested
                    if save_md:
                        md_filename = f"{os.path.splitext(file_name)[0]}.md"
                        with open(md_filename, 'w', encoding='utf-8') as f:
                            f.write(f"# Extracted Tables for {file_name}\n\n")
                            for i, res in enumerate(processed_results):
                                f.write(f"## Table {i+1}\n\n{res['md']}\n\n")
                        logger.info(f"\tMarkdown saved to: {md_filename}")
                        
                    # Save CSV if requested
                    if save_csv:
                        csv_filename = f"{os.path.splitext(file_name)[0]}.csv"
                        combined_df.to_csv(csv_filename, index=False, header=False)
                        logger.info(f"\tCSV saved to: {csv_filename}")

                    logger.info(f"\n\tExtraction complete for {file_name}. Saved to sheet: {sheet_name}")
                else:
                    logger.info(f"\n\tNo tables found in images for {file_name}.")
                    
        logger.info(f"Results saved to {excel_name}")
    except Exception as e:
        logger.error(f"\n\tProcessing error: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="\n\tExtract tables from PDF images and save to Excel/MD/CSV.")
    parser.add_argument("pdf_files", nargs="+", help="Paths to PDF files to process.")
    parser.add_argument("-o", "--output", default="extracted_tables.xlsx", help="Output Excel filename.")
    parser.add_argument("--md", action="store_true", help="Save raw markdown output.")
    parser.add_argument("--csv", action="store_true", help="Save as CSV file.")
    parser.add_argument("--clean", action="store_true", help="Clean/normalize data (remove currency symbols, standardize numbers).")
    
    args = parser.parse_args()
    main(args.pdf_files, args.output, save_md=args.md, save_csv=args.csv, clean=args.clean)

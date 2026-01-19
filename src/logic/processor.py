import os
import time
import re
import pandas as pd
from typing import List, Dict, Any, Optional
from google import genai
from PIL import Image

def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Cleans and normalizes DataFrame content."""
    def clean_cell(val):
        if pd.isna(val) or val is None: return ""
        s = str(val).strip()
        s = s.replace('$', '').replace('€', '').replace(',', '')
        try:
            if '.' in s: return float(s)
            return int(s)
        except ValueError:
            return s
    return df.apply(lambda col: col.map(clean_cell))

def parse_md(md_text: str) -> pd.DataFrame:
    """Parses Markdown table text into a pandas DataFrame."""
    lines = md_text.strip().split('\n')
    data = []
    for line in lines:
        if '|' in line:
            cells = [c.strip() for c in line.split('|')]
            if cells and cells[0] == '': cells = cells[1:]
            if cells and cells[-1] == '': cells = cells[:-1]
            if all(set(c) <= {'-', ':', ' '} for c in cells): continue
            if cells: data.append(cells)
        else:
            if line.strip(): data.append([line.strip()])
    
    if not data: return pd.DataFrame()
    
    max_cols = max(len(row) for row in data)
    norm_data = [row + [''] * (max_cols - len(row)) for row in data]
    return pd.DataFrame(norm_data)

ORDINAL_MAP = {
    "primera": 1, "primero": 1, "first": 1,
    "segunda": 2, "segundo": 2, "second": 2,
    "tercera": 3, "tercero": 3, "third": 3,
    "cuarta": 4, "cuarto": 4, "fourth": 4,
    "quinta": 5, "quinto": 5, "fifth": 5,
    "sexta": 6, "sexto": 6, "sixth": 6,
    "séptima": 7, "séptimo": 7, "seventh": 7,
    "octava": 8, "octavo": 8, "eighth": 8,
    "novena": 9, "noveno": 9, "ninth": 9,
    "décima": 10, "décimo": 10, "tenth": 10,
    "última": -1, "último": -1, "last": -1
}

def parse_page_query(prompt: str, total_pages: int, current_filename: str = None, all_filenames: List[str] = []) -> List[int]:
    """
    Parses the prompt to find page references. 
    Supports document-specific requests: "página 1 de DocA, página 2 de DocB".
    """
    prompt_lower = prompt.lower()
    
    # 0. Identify which files are mentioned in the entire prompt
    mentioned_files = []
    if all_filenames:
        for f in all_filenames:
            name_no_ext = os.path.splitext(f)[0].lower()
            pattern = rf"\b({re.escape(f.lower())}|{re.escape(name_no_ext)})\b"
            if re.search(pattern, prompt_lower):
                mentioned_files.append(f)
    
    # 1. Determine target text for the current file
    target_text = prompt_lower
    if mentioned_files:
        # If any files are mentioned but NOT the current one, skip this file
        if current_filename not in mentioned_files:
            return []
            
        # Split prompt into semantic blocks (by connectors and punctuation)
        # Delimiters: , ; . and y e (Spanish)
        blocks = re.split(r'[,;.]|\b(?:y|and|e)\b', prompt_lower)
        blocks = [b.strip() for b in blocks if b.strip()]
        
        # Find blocks that mention the current file
        name_no_ext = os.path.splitext(current_filename)[0].lower()
        cur_pattern = rf"\b({re.escape(current_filename.lower())}|{re.escape(name_no_ext)})\b"
        
        relevant_blocks = [b for b in blocks if re.search(cur_pattern, b)]
        
        if relevant_blocks:
            target_text = " ".join(relevant_blocks)
            # Check if these specific blocks contain any page instructions
            # We'll do a quick check to see if it's worth restricting to these blocks
            has_instr = any(re.search(r"\b\d+\b|p\u00e1gina|page|first|primera|last|\u00faltima", b) for b in relevant_blocks)
            if not has_instr:
                # Fallback: file was mentioned in a block without instructions (e.g. "p1 de A y B")
                # Use the full prompt for search but stay within "mentioned" mode
                target_text = prompt_lower
        else:
            # Fallback for complex nesting
            target_text = prompt_lower

    selected_pages = set()
    
    # 2. Check for ranges: "páginas 1 a 3", "pages 2-4", "p1-3"
    range_matches = re.finditer(r"(?:p\u00e1ginas?|pages?|p)\s*(\d+)\s*(?:a|to|-)\s*(\d+)", target_text)
    for match in range_matches:
        start = int(match.group(1))
        end = int(match.group(2))
        for p in range(start, end + 1):
            if 1 <= p <= total_pages:
                selected_pages.add(p - 1)

    # 3. Check for numeric single pages or lists: "página 2", "page 3", "páginas 1, 3, 5", "p1, p3"
    numeric_parts = re.split(r"(?:p\u00e1ginas?|pages?|p)", target_text)
    if len(numeric_parts) > 1:
        for part in numeric_parts[1:]:
            potential_numbers = re.findall(r"\b\d+\b", part)
            for num_str in potential_numbers:
                p = int(num_str)
                if 1 <= p <= total_pages:
                    selected_pages.add(p - 1)

    # 4. Check for ordinal words: "primera página", "last page"
    for word, val in ORDINAL_MAP.items():
        if re.search(rf"\b{word}\b", target_text):
            actual_p = val if val > 0 else total_pages + val + 1
            if 1 <= actual_p <= total_pages:
                selected_pages.add(actual_p - 1)
    
    if selected_pages:
        return sorted(list(selected_pages))

    # Fallback: if filename mentioned but no specific pages found, return all its pages.
    # Otherwise return all pages (global mode)
    return list(range(total_pages))

def extract_from_page(client: genai.Client, page: Any, prompt: str, log_callback=None, error_tracker: Dict[str, bool] = None) -> List[Dict[str, Any]]:
    """Extracts tables from a single PDF page."""
    if log_callback:
        log_callback(page.page_number)
        
    try:
        img = page.to_image(resolution=300).original
        
        max_retries = 1
        md_text = ""
        for attempt in range(max_retries):
            try:
                response = client.models.generate_content(
                    model='gemini-3-flash-preview',
                    contents=[prompt, img]
                )
                if response and hasattr(response, 'text') and response.text:
                    md_text = response.text
                break
            except Exception as e:
                if "400" in str(e):
                    if error_tracker is not None: error_tracker["has_error"] = True
                    raise e
                elif "403" in str(e):
                    if error_tracker is not None: error_tracker["has_error"] = True
                    raise e
                elif "429" in str(e):
                    if attempt < max_retries - 1:
                        time.sleep((attempt + 1) * 10)
                    else:
                        if error_tracker is not None: error_tracker["has_error"] = True
                        raise e
                else:
                    raise e
        
        clean_md = md_text.replace("```markdown", "").replace("```", "").strip()
        if clean_md:
            df = parse_md(clean_md)
            if not df.empty:
                return [{"df": df, "md": clean_md}]
    except Exception as e:
        if "400" in str(e) or "429" in str(e) or "403" in str(e):
            raise e
        else:
            raise e
    return []

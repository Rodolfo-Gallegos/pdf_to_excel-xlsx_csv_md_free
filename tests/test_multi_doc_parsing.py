import sys
import os

# Add src to path
sys.path.append('/home/rodolfo/Descargas/PDF_to_XLSX')

from src.logic.processor import parse_page_query

def test_multi_query(prompt, filenames, total_pages_map, expected_map):
    print(f"\nPrompt: '{prompt}'")
    all_pass = True
    for fname in filenames:
        tp = total_pages_map.get(fname, 5)
        expected = expected_map.get(fname)
        result = parse_page_query(prompt, tp, fname, filenames)
        status = "PASS" if result == expected else "FAIL"
        print(f"  - File: {fname:15} | Result: {result} | Expected: {expected} | {status}")
        if result != expected:
            all_pass = False
    return all_pass

if __name__ == "__main__":
    files = ["docA.pdf", "docB.pdf", "docC.pdf"]
    t_pages = {"docA.pdf": 3, "docB.pdf": 3, "docC.pdf": 3}
    
    scenarios = [
        # 1. Targeted pages for specific files
        {
            "prompt": "página 1 de docA y página 2 de docB",
            "expected": {"docA.pdf": [0], "docB.pdf": [1], "docC.pdf": []} # docC ignored because others were mentioned
        },
        # 2. Targeted ordinals
        {
            "prompt": "extrae la primera de docA y la última de docB",
            "expected": {"docA.pdf": [0], "docB.pdf": [2], "docC.pdf": []}
        },
        # 3. Global selection (none of the files mentioned)
        {
            "prompt": "página 2",
            "expected": {"docA.pdf": [1], "docB.pdf": [1], "docC.pdf": [1]}
        },
        # 4. Partial mention (one file mentioned without pages, others not mentioned)
        {
            "prompt": "procesar docA",
            "expected": {"docA.pdf": [0, 1, 2], "docB.pdf": [], "docC.pdf": []}
        },
        # 5. Mixed: ranges and ordinals
        {
            "prompt": "docA páginas 1 a 2, docB tercera página",
            "expected": {"docA.pdf": [0, 1], "docB.pdf": [2], "docC.pdf": []}
        },
        # 6. Matching by name without extension
        {
            "prompt": "docA página 1, docB página 3",
            "expected": {"docA.pdf": [0], "docB.pdf": [2], "docC.pdf": []}
        }
    ]
    
    overall_all_pass = True
    for s in scenarios:
        if not test_multi_query(s["prompt"], files, t_pages, s["expected"]):
            overall_all_pass = False
    
    if overall_all_pass:
        print("\nAll Multi-Document tests passed!")
    else:
        print("\nSome Multi-Document tests failed.")
        sys.exit(1)

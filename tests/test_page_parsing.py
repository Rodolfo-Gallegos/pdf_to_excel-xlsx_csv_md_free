import sys
import os

# Add src to path
sys.path.append('/home/rodolfo/Descargas/PDF_to_XLSX')

from src.logic.processor import parse_page_query

def test_query(prompt, total_pages, expected):
    result = parse_page_query(prompt, total_pages)
    status = "PASS" if result == expected else "FAIL"
    print(f"Prompt: '{prompt}' | Result: {result} | Expected: {expected} | {status}")
    return result == expected

if __name__ == "__main__":
    t_pages = 5
    tests = [
        ("Extrae la página 2", t_pages, [1]),
        ("Procesar páginas 1 a 3", t_pages, [0, 1, 2]),
        ("Solo la primera página", t_pages, [0]),
        ("The second page please", t_pages, [1]),
        ("La última página", t_pages, [4]),
        ("Extrae primero y segundo", t_pages, [0, 1]),
        ("Extract the matching tables", t_pages, [0, 1, 2, 3, 4]), # fallback
        ("Páginas 2 y 4", t_pages, [1, 3]), # should work via multiples if implemented, but currently regex is a bit simple. Wait, my new logic for ordinals handles multiple words!
    ]
    
    all_pass = True
    for p, tp, exp in tests:
        if not test_query(p, tp, exp):
            all_pass = False
    
    if all_pass:
        print("\nAll tests passed!")
    else:
        print("\nSome tests failed.")
        sys.exit(1)

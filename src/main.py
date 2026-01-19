import tkinter as tk
import os
from PIL import Image, ImageTk
from src.ui.app import PDFToXLSXGUI

def main():
    root = tk.Tk()
    
    # Set window icon
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_dir, "assets", "icons", "pdf_to_excel.png")
        if os.path.exists(icon_path):
            img = Image.open(icon_path)
            photo = ImageTk.PhotoImage(img)
            root.wm_iconphoto(True, photo)
    except Exception as e:
        print(f"Could not load icon: {e}")
        
    app = PDFToXLSXGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

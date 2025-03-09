from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENTATION
import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image

def select_images(title="Seleziona immagini"):
    """Apre il file system per selezionare immagini e restituisce i percorsi dei file selezionati."""
    root = tk.Toplevel()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title=title, filetypes=[("Immagini", "*.jpg;*.jpeg;*.png;*.bmp;*.gif")])
    root.destroy()
    return list(file_paths)

def filter_images_by_operator(image_paths):
    """Filtra le immagini in base all'operatore riconoscendo le sigle nel nome del file."""
    operators = {"Iliad": [], "TIM": [], "VF": [], "W3": []}
    for img in image_paths:
        for op in operators:
            if op.lower() in os.path.basename(img).lower():
                operators[op].append(img)
                break
    return operators

def create_document(output_path, all_blocks, operator_images):
    """Crea un unico documento Word combinando i blocchi con gli operatori."""
    doc = Document()
    
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(12)
    
    # Imposta l'orientamento della pagina in orizzontale
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    
    for block_title, block_images in all_blocks.items():
        for operator in ["TIM", "W3", "VF", "Iliad"]:
            filtered_images = [img for img in block_images if img in operator_images.get(operator, [])]
            if filtered_images:
                doc.add_heading(f"{block_title} - {operator}", level=2).paragraph_format.space_after = Inches(0.2)
                for image_path in filtered_images:
                    try:
                        doc.add_picture(image_path, width=Inches(8.5))
                    except Exception as e:
                        print(f"Errore nell'aggiunta dell'immagine {image_path}: {e}")
    
    file_path = os.path.abspath(output_path + ".docx")
    doc.save(file_path)
    print(f"Documento Word creato: {file_path}")

if __name__ == "__main__":
    document_title = input("Inserisci il titolo del documento: ")
    all_blocks = {}
    
    while True:
        if input("\nVuoi aggiungere un nuovo blocco di immagini? (s/n): ").strip().lower() not in ["s", "si", "yes", "y"]:
            break
        
        block_title = input("Inserisci il titolo per questo blocco di immagini: ")
        image_paths = select_images(f"Seleziona immagini per il blocco: {block_title}")
        
        if not image_paths:
            print("Nessuna immagine selezionata per questo blocco. Riprova.")
            continue
        
        all_blocks[block_title] = image_paths
    
    if all_blocks:
        operator_images = filter_images_by_operator(sum(all_blocks.values(), []))
        output_path = f"{document_title.replace(' ', '_')}_Unico"
        create_document(output_path, all_blocks, operator_images)
    else:
        print("Nessuna immagine selezionata. Uscita...")

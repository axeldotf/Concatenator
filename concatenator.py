from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENTATION
import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image
from tqdm import tqdm
from pathlib import Path

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
            if op.lower() in Path(img).stem.lower():
                operators[op].append(img)
                break
    return operators

def create_or_update_document(output_path, section_title, images):
    """Crea o aggiorna un documento Word con le immagini fornite."""
    file_path = Path(f"{output_path}.docx")
    
    if file_path.exists():
        file_path.unlink()
    
    doc = Document()
    
    # Imposta Arial 12 come font di default
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(12)
    
    # Imposta l'orientamento della pagina in orizzontale
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    
    max_width = (section.page_width - section.left_margin - section.right_margin) / 914400  # Conversione a pollici
    
    doc.add_heading(section_title, level=2).paragraph_format.space_after = Inches(0.2)
    
    # Mostra la barra di avanzamento
    for image_path in tqdm(images, desc=f"Aggiunta immagini a {section_title}", unit="img", ncols=80, leave=True):  
        try:
            with Image.open(image_path) as img:
                width, height = img.size
                aspect_ratio = height / width
                new_width = max_width
                new_height = new_width * aspect_ratio
            
            doc.add_picture(image_path, width=Inches(new_width), height=Inches(new_height))
        except Exception as e:
            print(f"Errore nell'aggiunta dell'immagine {image_path}: {e}")
    
    doc.save(file_path)
    return str(file_path)

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
        output_paths = set()
        
        for operator, images in operator_images.items():
            if images:
                output_path = f"{document_title.replace(' ', '_')}_{operator}"
                output_paths.add(Path(f"{output_path}.docx").resolve())
                for block_title, block_images in all_blocks.items():
                    filtered_images = [img for img in block_images if img in images]
                    if filtered_images:
                        create_or_update_document(output_path, block_title, filtered_images)
        
        print("\nDocumenti Word creati:")
        for path in output_paths:
            print(path)
    else:
        print("Nessuna immagine selezionata. Uscita...")

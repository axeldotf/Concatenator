from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENTATION
import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image
from tqdm import tqdm
from pathlib import Path

OPERATORS = ["Iliad", "TIM", "VF", "W3"]

def select_images(title="Seleziona immagini"):
    """Seleziona immagini tramite file dialog."""
    root = tk.Toplevel()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title=title, filetypes=[("Immagini", "*.jpg;*.jpeg;*.png;*.bmp;*.gif")])
    root.destroy()
    return list(file_paths)

def select_output_folder():
    """Permette di selezionare una cartella di destinazione per i file finali."""
    print("Seleziona la cartella di destinazione per i file finali...")
    root = tk.Toplevel()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Seleziona la cartella di destinazione")
    root.destroy()
    return Path(folder_selected) if folder_selected else Path.cwd()

def filter_images_by_operator(image_paths):
    """Filtra le immagini per operatore in base al nome del file."""
    return {op: [img for img in image_paths if op.lower() in Path(img).stem.lower()] for op in OPERATORS}

def crop_white_borders(image_path):
    """Ritaglia automaticamente i bordi bianchi laterali di un'immagine."""
    try:
        with Image.open(image_path) as img:
            img = img.convert("RGB")
            img_data = img.load()
            width, height = img.size
            
            left, right = 0, width
            for x in range(width):
                if any(img_data[x, y] != (255, 255, 255) for y in range(height)):
                    left = x
                    break
            for x in range(width - 1, -1, -1):
                if any(img_data[x, y] != (255, 255, 255) for y in range(height)):
                    right = x + 1
                    break

            if right > left:
                temp_path = image_path.replace(".", "_cropped.")
                img.crop((left, 0, right, height)).save(temp_path)
                return temp_path
    except Exception as e:
        print(f"Errore nel ritaglio di {image_path}: {e}")
    return image_path

def create_or_update_document(file_path, section_title, images, crop_images, block_count):
    """Crea o aggiorna un documento Word con immagini ritagliate e ridimensionate."""
    print(f"Elaborazione del blocco {block_count}: {section_title}")
    doc = Document(file_path) if file_path.exists() else Document()
    
    if file_path.exists():
        while len(doc.paragraphs) > 0 and doc.paragraphs[0].text.strip() == "":
            doc.paragraphs[0]._element.getparent().remove(doc.paragraphs[0]._element)
    else:
        style = doc.styles["Normal"]
        style.font.name = "Arial"
        style.font.size = Pt(12)
        section = doc.sections[0]
        section.orientation = WD_ORIENTATION.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        doc.save(file_path)

    new_section = doc.add_section()
    new_section.top_margin = new_section.bottom_margin = Inches(0)
    new_section.left_margin = new_section.right_margin = Inches(0)
    max_width = new_section.page_width / 914400  

    doc.add_heading(section_title, level=2).paragraph_format.space_after = Inches(0.2)

    for image_path in tqdm(images, desc=f"Aggiunta a {section_title}", unit="img", ncols=80, leave=True):  
        try:
            cropped_path = crop_white_borders(image_path) if crop_images else image_path
            with Image.open(cropped_path) as img:
                aspect_ratio = img.height / img.width
                doc.add_picture(cropped_path, width=Inches(max_width), height=Inches(max_width * aspect_ratio))
            if crop_images and cropped_path != image_path:
                os.remove(cropped_path)
        except Exception as e:
            print(f"Errore nell'aggiunta di {image_path}: {e}")

    doc.save(file_path)

if __name__ == "__main__":
    document_title = input("Inserisci il titolo del documento: ").replace(" ", "_")
    crop_option = input("Vuoi ritagliare automaticamente le immagini? (s/n): ").strip().lower()
    crop_images = crop_option == "s"
    output_folder = select_output_folder()
    
    output_paths = {op: output_folder / f"{document_title}_{op}_cut.docx" for op in OPERATORS}
    
    for file_path in output_paths.values():
        if file_path.exists():
            file_path.unlink()
    
    all_blocks = {}
    block_count = 0
    
    while input("\nAggiungere un blocco di immagini? (s/n): ").strip().lower() in ["s", "si", "yes", "y"]:
        block_count += 1
        print(f"Blocco {block_count} in corso...")
        block_title = input("Titolo del blocco: ")
        images = select_images(f"Seleziona immagini per: {block_title}")
        if images:
            all_blocks[block_title] = images
        else:
            print("Nessuna immagine selezionata. Riprova.")
    
    if all_blocks:
        operator_images = filter_images_by_operator(sum(all_blocks.values(), []))
        for operator, images in operator_images.items():
            if images:
                for block_count, (block_title, block_images) in enumerate(all_blocks.items(), start=1):
                    create_or_update_document(output_paths[operator], block_title, [img for img in block_images if img in images], crop_images, block_count)
    
        print("\nDocumenti creati nella cartella:", output_folder)
        for path in output_paths.values():
            print(path.resolve())
    else:
        print("Nessuna immagine selezionata. Uscita...")

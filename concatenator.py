from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENTATION
import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageDraw, ImageFont
from tqdm import tqdm
from pathlib import Path

OPERATORS = ["Iliad", "TIM", "VF", "W3"]
ORDER = [
    "LTE800 RSRP", "LTE800 QUAL", "GSM900 RXLEV", "UMTS900 RSCP", "UMTS900 QUAL", 
    "LTE1800 RSRP", "LTE1800 QUAL", "LTE2100 RSRP", "LTE2100 QUAL", "LTE2100 RSRP B100", 
    "LTE RSRQ B100", "UMTS2100 RSCP", "UMTS2100 QUAL", "LTE2600 RSRP", "LTE2600 QUAL", 
    "RSRP 3500", "RSRQ 3500"
]

def select_images(title="Seleziona immagini"):
    root = tk.Toplevel()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title=title, filetypes=[("Immagini", "*.jpg;*.jpeg;*.png;*.bmp;*.gif")])
    root.destroy()
    return list(file_paths)

def select_output_folder():
    print("Seleziona la cartella di destinazione per i file finali...")
    root = tk.Toplevel()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Seleziona la cartella di destinazione")
    root.destroy()
    return Path(folder_selected) if folder_selected else Path.cwd()

def filter_images_by_operator(image_paths):
    return {op: [img for img in image_paths if op.lower() in Path(img).stem.lower()] for op in OPERATORS}

def extract_label_name(image_path):
    filename = Path(image_path).stem
    for op in OPERATORS:
        if op in filename:
            tech_part = filename.replace("_Workbook_", "").replace(op, "").strip()
            return f"{op} {tech_part}"
    return filename

def crop_white_borders_and_add_label(image_path):
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
                cropped_img = img.crop((left, 0, right, height))
                
                label_text = extract_label_name(image_path)
                font = ImageFont.truetype("arial.ttf", 26)
                text_width, text_height = font.getbbox(label_text)[2:]
                
                padding_x = 20  # Spazio ai lati del testo
                padding_y = 10  # Spazio sopra e sotto il testo
                
                label_height = text_height + 2 * padding_y
                label_width = text_width + 2 * padding_x
                
                new_img = Image.new("RGB", (max(cropped_img.width, label_width), cropped_img.height + label_height), "white")
                new_img.paste(cropped_img, (0, label_height))
                
                draw = ImageDraw.Draw(new_img)
                draw.rectangle([(0, 0), (new_img.width, label_height)], outline="red", width=3)
                text_x = (new_img.width - text_width) // 2
                draw.text((text_x, padding_y), label_text, fill="black", font=font)
                
                temp_path = image_path.replace(".", "_labeled.")
                new_img.save(temp_path)
                return temp_path
    except Exception as e:
        print(f"Errore nel ritaglio di {image_path}: {e}")
    return image_path

def sort_images_by_order(images):
    def get_order_key(image_path):
        filename = Path(image_path).stem.lower()
        for index, label in enumerate(ORDER):
            if label.lower() in filename:
                return index
        return len(ORDER)  
    
    return sorted(images, key=get_order_key)

def create_or_update_document(file_path, section_title, images, crop_images, block_count):
    print(f"Elaborazione del blocco {block_count}: {section_title}")
    doc = Document(file_path) if file_path.exists() else Document()
    
    if not file_path.exists():
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
    
    sorted_images = sort_images_by_order(images)
    
    for image_path in tqdm(sorted_images, desc=f"Aggiunta a {section_title}", unit="img", ncols=80, leave=True):  
        try:
            processed_path = crop_white_borders_and_add_label(image_path) if crop_images else image_path
            with Image.open(processed_path) as img:
                aspect_ratio = img.height / img.width
                doc.add_picture(processed_path, width=Inches(max_width), height=Inches(max_width * aspect_ratio))
            if crop_images and processed_path != image_path:
                os.remove(processed_path)
        except Exception as e:
            print(f"Errore nell'aggiunta di {image_path}: {e}")
    
    doc.save(file_path)


if __name__ == "__main__":
    document_title = input("Inserisci il titolo del documento: ").replace(" ", "_")
    crop_option = input("Vuoi ritagliare automaticamente le immagini? (s/n): ").strip().lower()
    crop_images = crop_option in ["s", "si", "y", "yes"]
    output_folder = select_output_folder()
    
    output_paths = {op: output_folder / f"{document_title}_{op}_cut.docx" for op in OPERATORS}
    
    all_blocks = {}
    block_count = 0
    
    while input("\nAggiungere un blocco di immagini? (s/n): ").strip().lower() in ["s", "si", "yes", "y"]:
        block_count += 1
        block_title = input("Titolo del blocco: ")
        images = select_images(f"Seleziona immagini per: {block_title}")
        if images:
            all_blocks[block_title] = images
    
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

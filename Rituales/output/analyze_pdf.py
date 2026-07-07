#!/usr/bin/env python3
"""
Análisis detallado del PDF original para conversión fiel
"""
import fitz  # PyMuPDF
import json
from pathlib import Path

class PDFEncoder(json.JSONEncoder):
    """Custom JSON encoder for PyMuPDF objects"""
    def default(self, obj):
        if hasattr(obj, '__iter__'):
            try:
                return list(obj)
            except:
                return str(obj)
        return str(obj)

def analyze_pdf(pdf_path, output_json):
    """Analiza el PDF página por página y guarda el resultado"""
    doc = fitz.open(pdf_path)
    
    analysis = {
        "total_pages": len(doc),
        "metadata": doc.metadata,
        "pages": []
    }
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        page_info = {
            "page_number": page_num + 1,
            "width": page.rect.width,
            "height": page.rect.height,
            "text_blocks": [],
            "images": [],
            "links": [],
            "annotations": []
        }
        
        # Extraer bloques de texto con formato
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if "lines" in block:  # Bloque de texto
                block_text = ""
                block_font = None
                block_size = None
                block_bold = False
                
                for line in block["lines"]:
                    for span in line["spans"]:
                        block_text += span["text"]
                        block_font = span["font"]
                        block_size = span["size"]
                        block_bold = "Bold" in span["font"] or "bold" in span["font"]
                
                page_info["text_blocks"].append({
                    "text": block_text.strip(),
                    "font": block_font,
                    "size": block_size,
                    "bold": block_bold,
                    "bbox": list(block["bbox"])
                })
            
            elif "image" in block:  # Bloque de imagen
                page_info["images"].append({
                    "bbox": list(block["bbox"]),
                    "width": block["width"],
                    "height": block["height"]
                })
        
        # Extraer enlaces
        links = page.get_links()
        page_info["links"] = []
        for link in links:
            link_data = {}
            for k, v in link.items():
                if hasattr(v, '__iter__') and not isinstance(v, str):
                    try:
                        link_data[k] = list(v)
                    except:
                        link_data[k] = str(v)
                else:
                    link_data[k] = v
            page_info["links"].append(link_data)
        
        # Extraer anotaciones
        annots = page.annots()
        if annots:
            page_info["annotations"] = [str(annot) for annot in annots]
        
        analysis["pages"].append(page_info)
        
        if (page_num + 1) % 50 == 0:
            print(f"Analizadas {page_num + 1}/{len(doc)} páginas...")
    
    doc.close()
    
    # Guardar análisis
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(analysis, f, ensure_ascii=False, indent=2, cls=PDFEncoder)
    
    print(f"\nAnálisis completo guardado en: {output_json}")
    print(f"Total de páginas: {analysis['total_pages']}")
    
    return analysis

if __name__ == "__main__":
    pdf_path = r"C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\completo Ritual.pdf"
    output_json = r"C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\pdf_analysis.json"
    
    print("Analizando PDF original...")
    analysis = analyze_pdf(pdf_path, output_json)
    
    # Resumen
    total_images = sum(len(p["images"]) for p in analysis["pages"])
    total_text_blocks = sum(len(p["text_blocks"]) for p in analysis["pages"])
    
    print(f"\nResumen:")
    print(f"- Páginas: {analysis['total_pages']}")
    print(f"- Bloques de texto: {total_text_blocks}")
    print(f"- Imágenes: {total_images}")

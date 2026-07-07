#!/usr/bin/env python3
"""
Script para reemplazar el nombre de la Gran Logia en todos los documentos
"""
import os
import re
from pathlib import Path

# Definir los textos a reemplazar
OLD_TEXT = "Gran Logia de España de Menfis-Mizraim del Rito Antiguo y Primitivo de Menfis-Mizraim y de la Orden de los Ritos Unidos de Menfis & Mizraim – Ritos unificados por José GARIBALDI"
NEW_TEXT = "Gran Logia Simbólica del Rito Antiguo y Primitivo de Menfis Misraim de la República del Paraguay"

# También variantes con diferentes formatos
VARIANTS = [
    OLD_TEXT,
    OLD_TEXT.replace("–", "-"),  # Con guión normal en lugar de em-dash
    OLD_TEXT.replace("–", "—"),  # Con em-dash más largo
    OLD_TEXT.replace("GARIBALDI", "Garibaldi"),  # Con minúsculas
]

def replace_in_file(filepath):
    """Reemplazar texto en un archivo"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        replacements = 0
        
        for variant in VARIANTS:
            if variant in content:
                count = content.count(variant)
                content = content.replace(variant, NEW_TEXT)
                replacements += count
                print(f"  [+] Reemplazadas {count} ocurrencias de variante")
        
        if content != original_content:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            return replacements
        return 0
    except Exception as e:
        print(f"  [X] Error procesando {filepath}: {e}")
        return 0
    except Exception as e:
        print(f"  [X] Error procesando {filepath}: {e}")
        return 0

def main():
    output_dir = Path(r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output')
    
    # Archivos a procesar
    files_to_process = []
    
    # HTML files
    html_dir = output_dir / 'html'
    if html_dir.exists():
        files_to_process.extend(html_dir.glob('*.html'))
    
    # Markdown files
    md_dir = output_dir / 'md'
    if md_dir.exists():
        files_to_process.extend(md_dir.glob('*.md'))
    
    # JSON analysis file
    json_file = output_dir / 'pdf_analysis.json'
    if json_file.exists():
        files_to_process.append(json_file)
    
    print(f"Procesando {len(files_to_process)} archivos...")
    print(f"Texto antiguo: {OLD_TEXT}")
    print(f"Texto nuevo: {NEW_TEXT}")
    print()
    
    total_replacements = 0
    for filepath in files_to_process:
        print(f"Procesando: {filepath.name}")
        replacements = replace_in_file(filepath)
        if replacements > 0:
            print(f"  -> Total: {replacements} reemplazos")
            total_replacements += replacements
        else:
            print(f"  -> Sin cambios")
    
    print()
    print(f"[OK] Completado: {total_replacements} reemplazos totales en {len(files_to_process)} archivos")
    print()
    print("Ahora regenerando archivos DOCX...")

if __name__ == "__main__":
    main()

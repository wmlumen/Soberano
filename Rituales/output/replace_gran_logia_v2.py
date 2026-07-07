#!/usr/bin/env python3
"""
Script mejorado para reemplazar el nombre de la Gran Logia en todos los documentos
Usa expresiones regulares para manejar variaciones de formato
"""
import os
import re
from pathlib import Path

# Texto de reemplazo
NEW_TEXT = "Gran Logia Simbólica del Rito Antiguo y Primitivo de Menfis Misraim de la República del Paraguay"

# Patrones regex para encontrar el texto (maneja saltos de línea, espacios, etc.)
PATTERNS = [
    # Patrón principal con em-dash
    r'Gran\s+Logia\s+de\s+Espa[ñn]a\s+de\s+Menfis[-\s]+Mizraim\s+del\s+Rito\s+Antiguo\s+y\s+Primitivo\s+de\s+Menfis[-\s]+Mizraim\s+y\s+de\s+la\s+Orden\s+de\s+los\s+Ritos\s+Unidos\s+de\s+Menfis\s*[&y]\s*Mizraim\s*[–\-—]\s*Ritos\s+(?:unificados|reunificados)\s+por\s+Jos[ée]\s+GARIBALDI',
    
    # Patrón sin em-dash
    r'Gran\s+Logia\s+de\s+Espa[ñn]a\s+de\s+Menfis[-\s]+Mizraim\s+del\s+Rito\s+Antiguo\s+y\s+Primitivo\s+de\s+Menfis[-\s]+Mizraim\s+y\s+de\s+la\s+Orden\s+de\s+los\s+Ritos\s+Unidos\s+de\s+Menfis\s*[&y]\s*Mizraim\s+Ritos\s+(?:unificados|reunificados)\s+por\s+Jos[ée]\s+GARIBALDI',
    
    # Patrón con "de la del" (variante encontrada)
    r'Gran\s+Logia\s+de\s+Espa[ñn]a\s+de\s+Menfis[-\s]+Mizraim\s+de\s+la\s+del\s+Rito\s+Antiguo\s+y\s+Primitivo\s+de\s+Menfis[-\s]+Mizraim\s+y\s+de\s+la\s+Orden\s+de\s+los\s+Ritos\s+Unidos\s+de\s+Menfis\s*[&y]\s*Mizraim',
    
    # Patrón más simple (solo la primera parte)
    r'Gran\s+Logia\s+de\s+Espa[ñn]a\s+de\s+Menfis[-\s]+Mizraim\s+del\s+Rito\s+Antiguo\s+y\s+Primitivo\s+de\s+Menfis[-\s]+Mizraim\s+y\s+de\s+la\s+Orden\s+de\s+los\s+Ritos\s+Unidos\s+de\s+Menfis\s*[&y]\s*Mizraim',
]

def replace_in_file(filepath):
    """Reemplazar texto en un archivo usando regex"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        total_replacements = 0
        
        for i, pattern in enumerate(PATTERNS, 1):
            matches = re.findall(pattern, content, re.IGNORECASE)
            if matches:
                count = len(matches)
                content = re.sub(pattern, NEW_TEXT, content, flags=re.IGNORECASE)
                total_replacements += count
                print(f"  [+] Patrón {i}: {count} reemplazos")
        
        if content != original_content:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            return total_replacements
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
    
    print(f"Procesando {len(files_to_process)} archivos...")
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

if __name__ == "__main__":
    main()

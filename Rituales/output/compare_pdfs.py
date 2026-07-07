#!/usr/bin/env python3
"""
Comparación visual entre PDF original y PDF de control
"""
import fitz
from pathlib import Path
import json

def compare_pdfs(original_path, control_path, output_report):
    """Compara dos PDFs página por página"""
    
    doc_orig = fitz.open(original_path)
    doc_ctrl = fitz.open(control_path)
    
    report = {
        "original_pages": len(doc_orig),
        "control_pages": len(doc_ctrl),
        "page_count_match": len(doc_orig) == len(doc_ctrl),
        "differences": []
    }
    
    print(f"PDF Original: {len(doc_orig)} páginas")
    print(f"PDF Control: {len(doc_ctrl)} páginas")
    print(f"Coincidencia de páginas: {report['page_count_match']}")
    print()
    
    # Comparar página por página
    min_pages = min(len(doc_orig), len(doc_ctrl))
    
    for page_num in range(min_pages):
        page_orig = doc_orig[page_num]
        page_ctrl = doc_ctrl[page_num]
        
        page_diff = {
            "page": page_num + 1,
            "issues": []
        }
        
        # Comparar dimensiones
        if page_orig.rect != page_ctrl.rect:
            page_diff["issues"].append({
                "type": "dimensions",
                "original": f"{page_orig.rect.width:.1f}x{page_orig.rect.height:.1f}",
                "control": f"{page_ctrl.rect.width:.1f}x{page_ctrl.rect.height:.1f}"
            })
        
        # Comparar texto
        text_orig = page_orig.get_text("text").strip()
        text_ctrl = page_ctrl.get_text("text").strip()
        
        if text_orig != text_ctrl:
            # Calcular similitud simple
            orig_words = set(text_orig.split())
            ctrl_words = set(text_ctrl.split())
            
            if orig_words and ctrl_words:
                similarity = len(orig_words & ctrl_words) / max(len(orig_words), len(ctrl_words))
                page_diff["issues"].append({
                    "type": "text_content",
                    "similarity": f"{similarity:.2%}",
                    "original_length": len(text_orig),
                    "control_length": len(text_ctrl)
                })
        
        # Comparar imágenes
        images_orig = page_orig.get_images()
        images_ctrl = page_ctrl.get_images()
        
        if len(images_orig) != len(images_ctrl):
            page_diff["issues"].append({
                "type": "image_count",
                "original": len(images_orig),
                "control": len(images_ctrl)
            })
        
        # Si hay diferencias, agregar al reporte
        if page_diff["issues"]:
            report["differences"].append(page_diff)
            print(f"Página {page_num + 1}: {len(page_diff['issues'])} diferencia(s)")
            for issue in page_diff["issues"]:
                print(f"  - {issue['type']}: {issue}")
    
    doc_orig.close()
    doc_ctrl.close()
    
    # Resumen
    print(f"\n{'='*60}")
    print(f"RESUMEN DE COMPARACIÓN")
    print(f"{'='*60}")
    print(f"Páginas comparadas: {min_pages}")
    print(f"Páginas con diferencias: {len(report['differences'])}")
    print(f"Páginas idénticas: {min_pages - len(report['differences'])}")
    print(f"Tasa de coincidencia: {(min_pages - len(report['differences'])) / min_pages:.1%}")
    
    # Guardar reporte
    with open(output_report, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    
    print(f"\nReporte guardado en: {output_report}")
    
    return report

if __name__ == "__main__":
    original_pdf = r"C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\completo Ritual.pdf"
    control_pdf = r"C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\doc\Rituales_Simbolicos_Completo_control.pdf"
    output_report = r"C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\comparison_report.json"
    
    print("Comparando PDFs...")
    print(f"Original: {original_pdf}")
    print(f"Control: {control_pdf}")
    print()
    
    report = compare_pdfs(original_pdf, control_pdf, output_report)

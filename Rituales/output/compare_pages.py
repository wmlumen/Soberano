import os, re

base = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html'

# Compare page content between standalone and completo
pairs = [
    ('Ritual_1', 'Ritual_1_Instruccion_1er_Grado_Aprendiz.html', 68, 70),
    ('Ritual_2', 'Ritual_2_Instruccion_2o_Grado_Companero.html', 133, 134),
    ('Ritual_3', 'Ritual_3_Maestro.html', 137, 138),
]

for label, fname, pg_start, pg_end in pairs:
    path1 = os.path.join(base, fname)
    path2 = os.path.join(base, 'Rituales_Simbolicos_Completo.html')
    
    with open(path1, 'r', encoding='utf-8') as f:
        c1 = f.read()
    with open(path2, 'r', encoding='utf-8') as f:
        c2 = f.read()
    
    # Get section between page markers
    start1 = c1.find(f'P\u00c1GINA {pg_start}')
    end1 = c1.find(f'P\u00c1GINA {pg_end}', start1) if pg_end else len(c1)
    start2 = c2.find(f'P\u00c1GINA {pg_start}')
    end2 = c2.find(f'P\u00c1GINA {pg_end}', start2) if pg_end else len(c2)
    
    if start1 < 0:
        print(f'{label}: P\u00e1gina {pg_start} no encontrada en archivo individual')
        continue
    if start2 < 0:
        print(f'{label}: P\u00e1gina {pg_start} no encontrada en Completo')
        continue
    
    sec1 = c1[start1:end1] if end1 > 0 else c1[start1:]
    sec2 = c2[start2:end2] if end2 > 0 else c2[start2:]
    
    # Remove page markers for comparison
    clean1 = re.sub(r'<div class="page-marker[^>]*>.*?</div>', '', sec1)
    clean2 = re.sub(r'<div class="page-marker[^>]*>.*?</div>', '', sec2)
    
    # Normalize
    clean1 = re.sub(r'\s+', ' ', clean1).strip()
    clean2 = re.sub(r'\s+', ' ', clean2).strip()
    
    if clean1 == clean2:
        print(f'{label} pág.{pg_start}: ID\u00c9NTICO')
    else:
        print(f'{label} pág.{pg_start}: DIFERENTE')
        # Show where they differ
        words1 = clean1.split()
        words2 = clean2.split()
        min_len = min(len(words1), len(words2))
        for i in range(min_len):
            if words1[i] != words2[i]:
                print(f'  Diferencia en palabra {i}:')
                print(f'    Individual: ...{words1[max(0,i-3):i+4]}...')
                print(f'    Completo:   ...{words2[max(0,i-3):i+4]}...')
                break
        
        print(f'  Longitudes: Individual={len(clean1)}, Completo={len(clean2)}')
    print()

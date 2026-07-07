import os, re

base = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html'

# Verify all standalone docs have identical content to Completo for shared pages
pairs = [
    ('Ritual_1', 'Ritual_1_Instruccion_1er_Grado_Aprendiz.html', [68, 70, 71, 73, 74, 75]),
    ('Ritual_2', 'Ritual_2_Instruccion_2o_Grado_Companero.html', [133, 134, 135, 136]),
    ('Ritual_3', 'Ritual_3_Maestro.html', list(range(137, 178))),
]

completo_path = os.path.join(base, 'Rituales_Simbolicos_Completo.html')
with open(completo_path, 'r', encoding='utf-8') as f:
    completo = f.read()

all_ok = True
for label, fname, expected_pages in pairs:
    path = os.path.join(base, fname)
    with open(path, 'r', encoding='utf-8') as f:
        individual = f.read()
    
    for pg in expected_pages:
        # Extract the page section from both files (content between markers)
        pmarker = f'— PÁGINA {pg} —'
        
        # In individual file
        ind_start = individual.find(pmarker)
        if ind_start < 0:
            print(f'{label}: pág.{pg} no encontrada en archivo individual')
            all_ok = False
            continue
        
        # Find next page marker or end
        next_marker = None
        for np in expected_pages:
            if np > pg:
                next_marker = f'— PÁGINA {np} —'
                break
        ind_end = individual.find(next_marker, ind_start) if next_marker else len(individual)
        ind_content = individual[ind_start:ind_end] if ind_end > 0 else individual[ind_start:]
        
        # In completo file
        comp_start = completo.find(pmarker)
        if comp_start < 0:
            print(f'{label}: pág.{pg} no encontrada en Completo')
            all_ok = False
            continue
        
        # Find next marker in completo
        pg_idx_in_completo = None
        # Search forward for any page number marker
        comp_end = len(completo)
        for mp in range(pg + 1, 400):
            m = completo.find(f'— PÁGINA {mp} —', comp_start + len(pmarker))
            if m > 0:
                comp_end = m
                break
        comp_content = completo[comp_start:comp_end]
        
        # Clean both for comparison (remove the page marker div itself)
        pat = re.compile(r'<div class="page-marker[^>]*>— PÁGINA \d+ —</div>')
        ind_clean = pat.sub('', ind_content).strip()
        comp_clean = pat.sub('', comp_content).strip()
        
        # Normalize whitespace
        ind_clean = re.sub(r'\s+', ' ', ind_clean)
        comp_clean = re.sub(r'\s+', ' ', comp_clean)
        
        if ind_clean == comp_clean:
            print(f'{label} pag.{pg}: IGUAL (OK)')
        else:
            print(f'{label} pag.{pg}: DIFERENTE')
            # Simple comparison
            if len(ind_clean) != len(comp_clean):
                print(f'  Longitudes: ind={len(ind_clean)}, comp={len(comp_clean)}')
            # Find first diff
            min_len = min(len(ind_clean), len(comp_clean))
            for c in range(min_len):
                if ind_clean[c] != comp_clean[c]:
                    print(f'  Primer diff en posición {c}:')
                    print(f'    ind: ...{ind_clean[max(0,c-30):c+50]}...')
                    print(f'    comp:...{comp_clean[max(0,c-30):c+50]}...')
                    break
            all_ok = False

if all_ok:
    print('\n✓ TODAS LAS PÁGINAS SON IDÉNTICAS entre archivos individuales y Completo')
else:
    print('\n✗ Hay diferencias detectadas')

import os

base = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html'

files = [
    'Ritual_1_Instruccion_1er_Grado_Aprendiz.html',
    'Ritual_2_Instruccion_2o_Grado_Companero.html',
    'Ritual_3_Maestro.html',
    'Grandes_Constituciones.html',
    'Rituales_Simbolicos_Completo.html',
]

for fname in files:
    path = os.path.join(base, fname)
    with open(path, 'rb') as f:
        data = f.read()
    
    # Count page markers
    marker = b'PAGINA'
    count = data.count(marker)
    size_kb = len(data) / 1024
    print(f'{fname}: {size_kb:.0f} KB, {count} páginas')
    
    # Show first 3 page markers
    idx = 0
    shown = 0
    for _ in range(3):
        idx = data.find(marker, idx)
        if idx < 0:
            break
        start = max(0, idx - 10)
        end = min(len(data), idx + 40)
        snippet = data[start:end].decode('utf-8', errors='replace').strip()
        print(f'  -> {snippet}')
        idx += 1
        shown += 1
    print()

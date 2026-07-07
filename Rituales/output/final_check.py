import os, re

base = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html'
files = [
    'Ritual_1_Instruccion_1er_Grado_Aprendiz.html',
    'Ritual_2_Instruccion_2o_Grado_Companero.html',
    'Ritual_3_Maestro.html',
    'Grandes_Constituciones.html',
    'Rituales_Simbolicos_Completo.html',
]

accent_chars = [
    ('a', b'\xc3\xa1'),
    ('e', b'\xc3\xa9'),
    ('i', b'\xc3\xad'),
    ('o', b'\xc3\xb3'),
    ('u', b'\xc3\xba'),
    ('n', b'\xc3\xb1'),
    ('N', b'\xc3\x91'),
    ('?', b'\xc2\xbf'),
    ('!', b'\xc2\xa1'),
]

for fname in files:
    path = os.path.join(base, fname)
    with open(path, 'rb') as f:
        data = f.read()
    
    size_kb = len(data) / 1024
    pages = len(re.findall(b'P\xc3\x81GINA \d+', data))
    
    missing = []
    for label, bs in accent_chars:
        count = data.count(bs)
        if count == 0:
            missing.append(label)
    
    status = 'OK' if not missing else 'FALTAN: ' + ','.join(missing)
    print(f'{fname}: {size_kb:.0f}KB, {pages} pag - {status}')

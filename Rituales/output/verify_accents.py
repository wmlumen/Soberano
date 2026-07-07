import os

# Detailed verification of ALL generated files
base = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output'

groups = {
    'OCR (escaneados)': ['html/Ritual_1_Instruccion_del_1er_Grado_-_Aprendiz.html',
                         'html/Ritual_2_Instruccion_del_2o_Grado_-_Companero.html',
                         'html/Ritual_3_Ritual_e_Instruccion_del_3er_Grado_-_Maestro.html'],
    'Texto extraible': ['html/Grandes_Constituciones_y_Reglamentos_Generales.html',
                        'html/Rituales_Simbolicos_Completo.html']
}

accent_bytes = {
    '\u00e1': b'\xc3\xa1',  # a
    '\u00e9': b'\xc3\xa9',  # e
    '\u00ed': b'\xc3\xad',  # i
    '\u00f3': b'\xc3\xb3',  # o
    '\u00fa': b'\xc3\xba',  # u
    '\u00f1': b'\xc3\xb1',  # n
    '\u00bf': b'\xc2\xbf',  # inverted ?
}

for group_name, files in groups.items():
    print(f'=== {group_name} ===')
    for rel_path in files:
        path = os.path.join(base, rel_path)
        with open(path, 'rb') as f:
            data = f.read()
        size_kb = len(data) / 1024
        print(f'\n  {os.path.basename(rel_path)} ({size_kb:.0f} KB)')
        for char, byte_seq in accent_bytes.items():
            count = data.count(byte_seq)
            status = 'OK' if count > 0 else 'AUSENTE'
            print(f'    {char}: {count:>4} ({status})')
    print()

import os

files_to_check = [
    (r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html\Grandes_Constituciones_y_Reglamentos_Generales.html', 'HTML - Constituciones'),
    (r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html\Rituales_Simbolicos_Completo.html', 'HTML - Completo'),
    (r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\md\Grandes_Constituciones_y_Reglamentos_Generales.md', 'MD - Constituciones'),
    (r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\md\Rituales_Simbolicos_Completo.md', 'MD - Completo'),
]

chars_to_check = {
    'á': b'\xc3\xa1',
    'é': b'\xc3\xa9',
    'í': b'\xc3\xad',
    'ó': b'\xc3\xb3',
    'ú': b'\xc3\xba',
    'ñ': b'\xc3\xb1',
    'Ñ': b'\xc3\x91',
    'ü': b'\xc3\xbc',
    '¿': b'\xc2\xbf',
    '¡': b'\xc2\xa1',
}

for path, label in files_to_check:
    with open(path, 'rb') as f:
        data = f.read()
    
    print(f'=== {label} ({len(data)} bytes) ===')
    issues = []
    for char, byte_seq in chars_to_check.items():
        count = data.count(byte_seq)
        if count == 0:
            # Check if ASCII replacement happened
            # Try to find the character in the source text
            issues.append(f'  {char} (0x{byte_seq.hex()}): FALTANTE ({count} ocurrencias)')
        else:
            print(f'  {char}: {count} ocurrencias - OK')
    
    if issues:
        print('  PROBLEMAS:')
        for i in issues:
            print(i)
    print()

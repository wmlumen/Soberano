import os

md_path = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\md\Ritual_1_Instruccion_del_1er_Grado_-_Aprendiz.md'
with open(md_path, 'rb') as f:
    data = f.read()

# Find first question
idx = data.find(b'**Pregunta')
if idx >= 0:
    end = data.find(b'\n', idx + 50)
    content = data[idx:end] if end > 0 else data[idx:idx+80]
    print('First question in MD:', content)
    print('Hex:', content.hex())
    print('Decoded:', content.decode('utf-8'))

# Check if '¿' and accents are correct
idx_bf = data.find(b'\xc2\xbf')
if idx_bf >= 0:
    print(f'\nFound ¿ at byte {idx_bf}: context={data[idx_bf-5:idx_bf+15]}')

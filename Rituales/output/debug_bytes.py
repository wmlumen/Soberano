import os

html_path = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html\Ritual_1_Instruccion_del_1er_Grado_-_Aprendiz.html'
with open(html_path, 'rb') as f:
    data = f.read()

idx = data.find(b'class="question"')
if idx >= 0:
    start = data.find(b'>', idx) + 1
    end = data.find(b'</div>', start)
    content = data[start:end]
    print('Question content bytes:', content[:100])
    print('Decoded:', content.decode('utf-8')[:100])

# Also check around "iQu" in the generated file
idx = data.find(b'iQu' if b'iQu' in data else b'\xc2\xbfQu')
if idx >= 0:
    print(f'Found at {idx}: {data[idx:idx+30]}')
    print(f'Hex: {data[idx:idx+30].hex()}')

import os

# Check the actual content of the generated HTML
html_path = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html\Ritual_1_Instruccion_del_1er_Grado_-_Aprendiz.html'
with open(html_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Find first question
import re
for m in re.finditer(r'<div class="question">(.*?)</div>', content, re.DOTALL):
    text = re.sub(r'<[^>]+>', '', m.group(1)).strip()
    print('First Q:', repr(text[:100]))
    print('First Q readable:', text[:100])
    break

for m in re.finditer(r'<div class="answer">(.*?)</div>', content, re.DOTALL):
    text = re.sub(r'<[^>]+>', '', m.group(1)).strip()
    print('First A:', repr(text[:100]))
    print('First A readable:', text[:100])
    break

# Check MD
md_path = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\md\Ritual_1_Instruccion_del_1er_Grado_-_Aprendiz.md'
with open(md_path, 'r', encoding='utf-8') as f:
    content = f.read()

lines = content.split('\n')
for line in lines:
    if 'Pregunta' in line and '**' in line:
        print('\nMD first question:', repr(line[:120]))
        print('MD first question readable:', line[:120])
        break

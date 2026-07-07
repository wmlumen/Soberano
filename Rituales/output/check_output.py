import re
import os

html_dir = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\html'

for fname in sorted(os.listdir(html_dir)):
    if not fname.endswith('.html'):
        continue
    path = os.path.join(html_dir, fname)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find all question divs
    qs = []
    for m in re.finditer(r'<div class="question">(.*?)</div>', content, re.DOTALL):
        text = re.sub(r'<[^>]+>', '', m.group(1)).strip()
        qs.append(text)
    
    # Find all answer divs
    ans = []
    for m in re.finditer(r'<div class="answer">(.*?)</div>', content, re.DOTALL):
        text = re.sub(r'<[^>]+>', '', m.group(1)).strip()
        ans.append(text)
    
    print(f'=== {fname} ===')
    print(f'  Questions: {len(qs)}')
    print(f'  Answers: {len(ans)}')
    if qs:
        print(f'  First Q: {qs[0][:80]}')
    if ans:
        print(f'  First A: {ans[0][:80]}')
    print()

# Check MD files
md_dir = r'C:\Temp\kitian\GITHUT\RItoMemphisMisraim\Rituales\output\md'
for fname in sorted(os.listdir(md_dir)):
    if not fname.endswith('.md'):
        continue
    path = os.path.join(md_dir, fname)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    q_lines = [l for l in lines if 'Pregunta:' in l]
    print(f'=== {fname} ({len(lines)} lines, {len(q_lines)} questions) ===')
    if q_lines:
        print(f'  First: {q_lines[0][:100]}')
    print()

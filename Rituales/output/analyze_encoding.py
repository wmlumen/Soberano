import os

path = os.path.join(os.environ['TEMP'], 'pdf_imgs', 'ritual_1_completo.txt')
with open(path, 'rb') as f:
    data = f.read()

# Count replacement characters
replacement_count = data.count(b'\xef\xbf\xbd')
print(f'Replacement chars (U+FFFD): {replacement_count}')

# Show first replacement
idx = data.find(b'\xef\xbf\xbd')
if idx >= 0:
    before = data[max(0,idx-20):idx]
    after = data[idx+3:idx+23]
    print(f'First replacement at byte {idx}:')
    print(f'  Before: {before}')
    print(f'  After:  {after}')

# Find all unique byte sequences that look like they should be accents
# The OCR likely output single-byte chars for what should be multi-byte utf-8
print('\nLooking for specific patterns...')
patterns = [b'\xed', b'\xe9', b'\xe1', b'\xf3', b'\xfa']
for p in patterns:
    count = data.count(p)
    if count > 0:
        idx = data.find(p)
        print(f'Byte 0x{p.hex()}: count={count}, first at {idx}: context={data[max(0,idx-5):idx+5]}')

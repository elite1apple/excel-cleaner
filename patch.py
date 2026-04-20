path = r'c:\Users\mdash\Downloads\Excel-clean\static\script.js'
with open(path, 'r', encoding='utf-8') as f:
    code = f.read()

old = "    const fc = document.getElementById('fabricColors').value.trim();"
new = "    const fc = collectFabricColors();"

if old in code:
    code = code.replace(old, new)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(code)
    print('Patched')
else:
    print('Target not found — searching context...')
    idx = code.find('fabricColors')
    print(repr(code[max(0,idx-10):idx+80]))

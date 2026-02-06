import shutil
shutil.copy('create_standard_template.py', 'create_new_template.py')

# Read and modify
with open('create_new_template.py', 'r') as f:
    content = f.read()

content = content.replace("'STANDARD_TEMPLATE.xlsx'", "'STANDARD_TEMPLATE_NEW.xlsx'")

with open('create_new_template.py', 'w') as f:
    f.write(content)

print("Created create_new_template.py")

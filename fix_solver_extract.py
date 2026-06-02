import re

with open('macros.vba', 'r') as f:
    content = f.read()

def extract_module(content, module_name):
    pattern = r"VBA MACRO " + re.escape(module_name) + r" \n.*?^- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - \n(.*?)(?=\n-{79}\n|\Z)"
    match = re.search(pattern, content, re.DOTALL | re.MULTILINE)
    if match:
        return match.group(1).strip()
    return ""

solver = extract_module(content, 'solverUtils.bas')

with open('solverUtils.bas', 'w') as f:
    f.write('Attribute VB_Name = "solverUtils"\n' + solver)

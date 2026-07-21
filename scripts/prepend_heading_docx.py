import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import shutil
import tempfile

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', ns['w'])

src = Path('Enqiang Linghu_keynote.docx')
backup = Path('Enqiang Linghu_keynote.docx.bak')

if not backup.exists():
    shutil.copy(src, backup)

with zipfile.ZipFile(src, 'r') as zin:
    items = {name: zin.read(name) for name in zin.namelist()}

xml = items['word/document.xml']
root = ET.fromstring(xml)
body = root.find('w:body', ns)

# Build new paragraph w:p with style Heading1
p = ET.Element('{%s}p' % ns['w'])
pp = ET.SubElement(p, '{%s}pPr' % ns['w'])
pstyle = ET.SubElement(pp, '{%s}pStyle' % ns['w'])
pstyle.set('{%s}val' % ns['w'], 'Heading1')
run = ET.SubElement(p, '{%s}r' % ns['w'])
t = ET.SubElement(run, '{%s}t' % ns['w'])
t.text = 'Plenary Talk 1: Enqiang Linghu'

# Insert paragraph as first element after possible sectPr or title
# Find index to insert before first existing paragraph element
insert_index = 0
for i, child in enumerate(list(body)):
    # skip if child is sectPr (section properties) at end
    insert_index = i
    break

body.insert(0, p)

# Serialize modified document.xml
new_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)

# Write new docx to temp file then replace original
with tempfile.NamedTemporaryFile(delete=False) as tmpf:
    tmpname = tmpf.name
with zipfile.ZipFile(tmpname, 'w') as zout:
    for name, data in items.items():
        if name == 'word/document.xml':
            zout.writestr(name, new_xml)
        else:
            zout.writestr(name, data)

shutil.move(tmpname, src)
print('Inserted heading into', src)

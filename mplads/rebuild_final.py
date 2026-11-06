"""
Proper rebuild: start from ORIGINAL template, copy slides 1-4 from edited version,
rebuild slides 5-6 from scratch with clean layout.
"""
import zipfile
import shutil
import os
import re
import sys
from copy import deepcopy

sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\SIH2026-IDEA-Presentation-Format.pptx'
EDITED   = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\ppt_v1.pptx'
WORK     = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\clean_final'
OUT      = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\ppt_v1_final.pptx'

# Step 1: Extract original template
if os.path.exists(WORK):
    shutil.rmtree(WORK)
os.makedirs(WORK)

with zipfile.ZipFile(TEMPLATE, 'r') as z:
    z.extractall(WORK)
print("Extracted original template")

# Step 2: Copy slides 1-4 from edited version
with zipfile.ZipFile(EDITED, 'r') as z:
    for sn in [1, 2, 3, 4]:
        data = z.read('ppt/slides/slide' + str(sn) + '.xml')
        path = os.path.join(WORK, 'ppt', 'slides', 'slide' + str(sn) + '.xml')
        with open(path, 'wb') as f:
            f.write(data)
        rels = z.read('ppt/slides/_rels/slide' + str(sn) + '.xml.rels')
        rels_path = os.path.join(WORK, 'ppt', 'slides', '_rels', 'slide' + str(sn) + '.xml.rels')
        with open(rels_path, 'wb') as f:
            f.write(rels)
print("Copied slides 1-4 from edited version")

# Step 3: Verify template slides 5-6 are clean
for sn in [5, 6]:
    path = os.path.join(WORK, 'ppt', 'slides', 'slide' + str(sn) + '.xml')
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    runs = re.findall(r'<a:t>([^<]+)</a:t>', content)
    print("Template slide " + str(sn) + " text:", [t for t in runs if len(t.strip()) > 1])

# Step 4: Pack into PPTX
import subprocess
# Use python to create zip
import zipfile as zf

# First, create the PPTX from the work directory
with zf.ZipFile(OUT, 'w', zf.ZIP_DEFLATED) as zout:
    for root, dirs, files in os.walk(WORK):
        for file in files:
            filepath = os.path.join(root, file)
            arcname = os.path.relpath(filepath, WORK)
            zout.write(filepath, arcname)

print("Packed to:", OUT)

# Verify
with zf.ZipFile(OUT, 'r') as z:
    slides = [n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
    print("Total slides:", len(slides))

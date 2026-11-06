import zipfile
import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

pptx_path = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\ppt_v1_edited.pptx'

with zipfile.ZipFile(pptx_path, 'r') as z:
    # Check slide 6 rels
    with z.open('ppt/slides/_rels/slide6.xml.rels') as f:
        content = f.read().decode('utf-8')
        print('=== Slide 6 Relationships ===')
        print(content[:3000])

    # Check slide 6 XML for hyperlink elements
    with z.open('ppt/slides/slide6.xml') as f:
        content = f.read().decode('utf-8')
        hlink = re.findall(r'hlinkClick', content)
        print('hlinkClick count:', len(hlink))

        # Find rIds
        rids = re.findall(r'r:id="(rId\d+)"', content)
        print('rIds in slide6:', set(rids))

        # Check presentation.xml.rels for hyperlink relationships
    with z.open('ppt/_rels/presentation.xml.rels') as f:
        rels = f.read().decode('utf-8')
        hyperlinks = re.findall(r'Type="[^"]*hyperlink"', rels)
        print('Hyperlink rels in presentation:', len(hyperlinks))

    # Check slide6 rels for hyperlink type
    with z.open('ppt/slides/_rels/slide6.xml.rels') as f:
        rels = f.read().decode('utf-8')
        hyperlinks = re.findall(r'Type="[^"]*hyperlink"', rels)
        print('Hyperlink rels in slide6:', len(hyperlinks))

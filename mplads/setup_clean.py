import zipfile, os, shutil, re, sys
sys.stdout.reconfigure(encoding='utf-8')

template = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\SIH2026-IDEA-Presentation-Format.pptx'
work = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\clean_v3'
if os.path.exists(work):
    shutil.rmtree(work)
os.makedirs(work)

with zipfile.ZipFile(template, 'r') as z:
    z.extractall(work)

src = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\ppt_v1.pptx'
with zipfile.ZipFile(src, 'r') as z:
    for sn in [1,2,3,4]:
        data = z.read('ppt/slides/slide' + str(sn) + '.xml')
        with open(os.path.join(work, 'ppt', 'slides', 'slide' + str(sn) + '.xml'), 'wb') as f:
            f.write(data)
        rels_data = z.read('ppt/slides/_rels/slide' + str(sn) + '.xml.rels')
        with open(os.path.join(work, 'ppt', 'slides', '_rels', 'slide' + str(sn) + '.xml.rels'), 'wb') as f:
            f.write(rels_data)

for sn in [5, 6]:
    path = os.path.join(work, 'ppt', 'slides', 'slide' + str(sn) + '.xml')
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    runs = re.findall(r'<a:t>([^<]+)</a:t>', content)
    print('Slide ' + str(sn) + ' text:')
    for t in runs:
        if len(t.strip()) > 1:
            print('  ' + t[:80])
    print()

import zipfile, re, sys
sys.stdout.reconfigure(encoding='utf-8')

# Check original template shapes
template = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\SIH2026-IDEA-Presentation-Format.pptx'
with zipfile.ZipFile(template, 'r') as z:
    for sn in [5, 6]:
        fname = 'ppt/slides/slide' + str(sn) + '.xml'
        with z.open(fname) as f:
            content = f.read().decode('utf-8')
            names = re.findall(r'name="([^"]+)"', content)
            print('Template slide ' + str(sn) + ' shape names:')
            for n in names:
                print('  ' + n)
            print()

# Check v2 file shapes
v2 = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\ppt_v1_v2.pptx'
with zipfile.ZipFile(v2, 'r') as z:
    for sn in [5, 6]:
        fname = 'ppt/slides/slide' + str(sn) + '.xml'
        with z.open(fname) as f:
            content = f.read().decode('utf-8')
            names = re.findall(r'name="([^"]+)"', content)
            print('V2 slide ' + str(sn) + ' shape names:')
            for n in names:
                print('  ' + n)
            print()

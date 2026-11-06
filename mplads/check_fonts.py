import zipfile
import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

pptx_path = r'C:\Users\Vaibhav\projects\hackathon_projects\sih_2026_bs\mplads\ppt_v1.pptx'

with zipfile.ZipFile(pptx_path, 'r') as z:
    fname = 'ppt/slides/slide4.xml'
    with z.open(fname) as f:
        content = f.read().decode('utf-8')

        # Find the section header style (Wingdings bullet + underlined text)
        runs = re.findall(r'<a:r>(.*?)</a:r>', content, re.DOTALL)
        print('=== Slide 4 all runs ===')
        for run in runs:
            t_match = re.search(r'<a:t>([^<]+)</a:t>', run)
            f_match = re.search(r'typeface="([^"]+)"', run)
            s_match = re.search(r'sz="(\d+)"', run)
            b_match = re.search(r' b="1"', run)
            u_match = re.search(r' u="([^"]+)"', run)
            if t_match:
                txt = t_match.group(1)[:50]
                face = f_match.group(1) if f_match else 'default'
                sz = int(s_match.group(1)) // 100 if s_match else 0
                bold = 'B' if b_match else ''
                u = 'U' if u_match else ''
                print('  ' + face + ' ' + str(sz) + 'pt ' + bold + u + ' "' + txt + '"')

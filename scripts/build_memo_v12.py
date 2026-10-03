"""Build the v12 memo: plain-language text from memo_text.py, bold section labels only,
four charts floated right, scenario table tied to the model.

Usage: python3 build_memo_v12.py <unzipped NCLH_memo_v10_reference.docx dir> <output dir>
then zip the output dir back into a .docx."""
import re, os, shutil, sys
from xml.sax.saxutils import escape
sys.path.insert(0, os.path.dirname(__file__))
import memo_text as M

src, out = sys.argv[1], sys.argv[2]
page_break_before = set(sys.argv[3].split(',')) if len(sys.argv) > 3 and sys.argv[3] else set()
if os.path.exists(out): shutil.rmtree(out)
shutil.copytree(src, out)
docp = os.path.join(out, 'word/document.xml')
x = open(docp, encoding='utf-8').read()

W_EMU, GAP_EMU = 2286000, 137160   # 2.5in charts
DROP = {'rId7': 'image2.png', 'rId9': 'image4.png', 'rId11': 'image6.png', 'rId13': 'image8.png'}

# --- pieces reused from the original file
drawings = {re.search(r'r:embed="(rId\d+)"', d).group(1): d
            for d in re.findall(r'<w:drawing>.*?</w:drawing>', x, re.S)}
i = x.index('Valuation: Debt'); j = x.index('<w:tbl>', i); k = x.index('</w:tbl>', j) + 8
table = x[j:k]
for old, new in M.TABLE_FIX.items():
    assert table.count(f'>{old}<') == 1, old
    table = table.replace(f'>{old}<', f'>{new}<')
sect = re.search(r'<w:sectPr.*?</w:sectPr>', x, re.S).group(0)
head = x[:x.index('<w:body>') + len('<w:body>')]

def anchor(rid, n):
    d = drawings[rid]
    cx, cy = map(int, re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', d).groups())
    ncy = round(cy * W_EMU / cx)
    graphic = re.search(r'<wp:cNvGraphicFramePr>.*?</a:graphic>', d, re.S).group(0)
    graphic = graphic.replace(f'cx="{cx}" cy="{cy}"', f'cx="{W_EMU}" cy="{ncy}"')
    name = re.search(r'<wp:docPr [^>]*name="([^"]*)"', d).group(1)
    return ('<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
            f'<wp:anchor distT="0" distB="45720" distL="{GAP_EMU}" distR="0" simplePos="0" '
            f'relativeHeight="{251659264 + n}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{W_EMU}" cy="{ncy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapSquare wrapText="left"/>'
            f'<wp:docPr id="{100 + n}" name="{name}"/>'
            f'{graphic}</wp:anchor></w:drawing></w:r>')

def run(text, bold=False):
    rpr = '<w:rPr><w:b/></w:rPr>' if bold else ''
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def para(inner, before=0, keep=False, brk=False):
    ppr = ('<w:pPr>' + ('<w:keepLines/>' if keep else '') + ('<w:pageBreakBefore/>' if brk else '')
           + (f'<w:spacing w:before="{before}" w:after="60"/>' if before else '<w:spacing w:after="60"/>')
           + '<w:jc w:val="both"/></w:pPr>')
    return f'<w:p>{ppr}{inner}</w:p>'

SECTION_GAP = 80   # twips of space above each section after the first
body, n = [], 0
def section(label, paras, rid=None, first=False):
    global n
    brk = label.split(':')[0] in page_break_before
    lead = ''
    if rid:
        n += 1
        lead = anchor(rid, n)
    body.append(para(lead + run(label + ' ', True) + run(paras[0]),
                     before=0 if (first or brk) else SECTION_GAP, keep=bool(rid), brk=brk))
    for p in paras[1:]:
        body.append(para(run(p)))

for idx, (label, paras, rid) in enumerate(M.SECTIONS):
    section(label, paras, rid, first=(idx == 0))
    if label == M.TABLE_AFTER:
        body.append(table)
for label, paras in M.AFTER_TABLE:
    section(label, paras)
body.append('<w:p><w:pPr><w:spacing w:before="60" w:after="0"/><w:jc w:val="both"/></w:pPr>'
            f'<w:r><w:rPr><w:color w:val="404040"/><w:sz w:val="13"/></w:rPr><w:t xml:space="preserve">{escape(M.SOURCES)}</w:t></w:r></w:p>')

x = head + ''.join(body) + sect + '</w:body></w:document>'
open(docp, 'w', encoding='utf-8').write(x)

relp = os.path.join(out, 'word/_rels/document.xml.rels')
r = open(relp, encoding='utf-8').read()
for rid, img in DROP.items():
    r = re.sub(rf'<Relationship Id="{rid}"[^>]*/>', '', r)
    os.remove(os.path.join(out, 'word/media', img))
open(relp, 'w', encoding='utf-8').write(r)
print('built', out)

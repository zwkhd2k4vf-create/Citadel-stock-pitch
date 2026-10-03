"""Re-layout NCLH memo: one essential chart per section, floated right with text wrap."""
import re, shutil, os, sys
src, out = sys.argv[1], sys.argv[2]
if os.path.exists(out): shutil.rmtree(out)
shutil.copytree(src, out)
docp = os.path.join(out, 'word/document.xml')
x = open(docp, encoding='utf-8').read()

KEEP = {'rId6', 'rId8', 'rId10', 'rId12'}   # price-implied EBITDA, Caribbean cuts, price guarantee, EBITDA bridge
DROP = {'rId7': 'image2.png', 'rId9': 'image4.png', 'rId11': 'image6.png', 'rId13': 'image8.png'}
W_EMU = 2148840            # 2.35in chart width
GAP_EMU = 137160           # 0.15in gap between text and chart

def top_level_tables(s):
    """Yield (start, end) of top-level <w:tbl> elements in body (handles nesting)."""
    i = 0
    while True:
        a = s.find('<w:tbl>', i)
        if a < 0: return
        depth, j = 0, a
        while True:
            o = s.find('<w:tbl>', j); c = s.find('</w:tbl>', j)
            if o != -1 and o < c: depth += 1; j = o + 7
            else:
                depth -= 1; j = c + 8
                if depth == 0: break
        yield a, j
        i = j

def split_cells(tbl):
    """Return list of inner xml of top-level cells of the single row."""
    cells, i = [], 0
    inner = tbl[tbl.index('<w:tr'):]
    while True:
        a = inner.find('<w:tc>', i)
        if a < 0: break
        depth, j = 0, a
        while True:
            o = inner.find('<w:tc>', j); c = inner.find('</w:tc>', j)
            if o != -1 and o < c: depth += 1; j = o + 6
            else:
                depth -= 1; j = c + 7
                if depth == 0: break
        cells.append(inner[a + 6:j - 7])
        i = j
        if len(cells) == 2: break
    return cells

def to_anchor(drawing, docpr_id, z):
    m = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', drawing)
    cx, cy = int(m.group(1)), int(m.group(2))
    ncy = round(cy * W_EMU / cx)
    graphic = re.search(r'<wp:cNvGraphicFramePr>.*?</a:graphic>', drawing, re.S).group(0)
    graphic = graphic.replace(f'cx="{cx}" cy="{cy}"', f'cx="{W_EMU}" cy="{ncy}"')
    name = re.search(r'<wp:docPr [^>]*name="([^"]*)"', drawing).group(1)
    return ('<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
            f'<wp:anchor distT="0" distB="45720" distL="{GAP_EMU}" distR="0" simplePos="0" '
            f'relativeHeight="{z}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{W_EMU}" cy="{ncy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapSquare wrapText="left"/>'
            f'<wp:docPr id="{docpr_id}" name="{name}"/>'
            f'{graphic}</wp:anchor></w:drawing></w:r>')

pieces, last, n = [], 0, 0
for a, b in list(top_level_tables(x)):
    tbl = x[a:b]
    left, right = split_cells(tbl)
    left = re.sub(r'^<w:tcPr>.*?</w:tcPr>', '', left, flags=re.S)
    drawings = re.findall(r'<w:drawing>.*?</w:drawing>', right, re.S)
    keep = [d for d in drawings if re.search(r'r:embed="(rId\d+)"', d).group(1) in KEEP]
    assert len(keep) == 1, keep
    n += 1
    anchor = to_anchor(keep[0], 100 + n, 251659264 + n)
    # insert anchor run at the start of the section's first paragraph (after its pPr)
    # keep each chart's lead paragraph on one page; Thesis 2 opens page 2 so its chart sits beside it
    extra = '<w:keepLines/>' + ('<w:pageBreakBefore/>' if 'Thesis 2:' in left[:3000] else '')
    left = left.replace('<w:pPr>', '<w:pPr>' + extra, 1)
    p_end = left.index('</w:pPr>') + len('</w:pPr>')
    left = left[:p_end] + anchor + left[p_end:]
    pieces += [x[last:a], left]
    last = b
pieces.append(x[last:])
x = ''.join(pieces)
assert '<wp:inline' not in x and x.count('<wp:anchor') == 4
open(docp, 'w', encoding='utf-8').write(x)

# drop the unused images and their relationships
relp = os.path.join(out, 'word/_rels/document.xml.rels')
r = open(relp, encoding='utf-8').read()
for rid, img in DROP.items():
    r = re.sub(rf'<Relationship Id="{rid}"[^>]*/>', '', r)
    os.remove(os.path.join(out, 'word/media', img))
open(relp, 'w', encoding='utf-8').write(r)
print('ok')

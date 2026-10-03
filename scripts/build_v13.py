"""Build the v13 memo from the v10 docx shell: plain-language text (memo_text_v13.py), new exhibits from the v6 model,
a split 'From the data to the model' section on page 2, and the scenario table.

usage: python3 build_v13.py <unzipped v10 docx dir> <out dir> <charts dir>"""
import os
import re
import shutil
import struct
import sys
from xml.sax.saxutils import escape

sys.path.insert(0, os.path.dirname(__file__))
import memo_text_v13 as M

src, out, charts = sys.argv[1:4]
if os.path.exists(out):
    shutil.rmtree(out)
shutil.copytree(src, out)
docp = os.path.join(out, 'word/document.xml')
x = open(docp, encoding='utf-8').read()

# new exhibits replace the four kept images; the other four are dropped
CHART = {'ex1': ('rId6', 'image1.png', 'ex1_ebitda.png'), 'ex2': ('rId8', 'image3.png', 'ex2_fares.png'),
         'ex3': ('rId10', 'image5.png', 'ex3_deposits.png'), 'ex4': ('rId12', 'image7.png', 'ex4_bridge.png')}
for rid, img, png in CHART.values():
    shutil.copy(os.path.join(charts, png), os.path.join(out, 'word/media', img))
relp = os.path.join(out, 'word/_rels/document.xml.rels')
r = open(relp, encoding='utf-8').read()
for rid, img in {'rId7': 'image2.png', 'rId9': 'image4.png', 'rId11': 'image6.png', 'rId13': 'image8.png'}.items():
    r = re.sub(rf'<Relationship Id="{rid}"[^>]*/>', '', r)
    os.remove(os.path.join(out, 'word/media', img))
open(relp, 'w', encoding='utf-8').write(r)

EMU = 914400
FLOAT_W = int(2.75 * EMU)
GAP = 137160
sect = re.search(r'<w:sectPr.*?</w:sectPr>', x, re.S).group(0)
head = x[:x.index('<w:body>') + len('<w:body>')]


def png_ratio(key):
    with open(os.path.join(charts, CHART[key][2]), 'rb') as f:
        f.read(16)
        w, h = struct.unpack('>II', f.read(8))
    return h / w


def graphic(key, cx, cy, n):
    rid = CHART[key][0]
    return ('<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
            '</wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="0" name="{key}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')


WIDTHS = {'ex1': int(2.45 * EMU), 'ex2': int(2.8 * EMU), 'ex3': int(2.8 * EMU)}


def anchor(key, n):
    fw = WIDTHS.get(key, FLOAT_W)
    cy = round(fw * png_ratio(key))
    return ('<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
            f'<wp:anchor distT="0" distB="45720" distL="{GAP}" distR="0" simplePos="0" relativeHeight="{251659264 + n}" '
            'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0"><wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{fw}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="left"/>'
            f'<wp:docPr id="{100 + n}" name="{key}"/>{graphic(key, fw, cy, n)}</wp:anchor></w:drawing></w:r>')


def inline(key, width_in, n):
    cx = int(width_in * EMU)
    cy = round(cx * png_ratio(key))
    return ('<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{100 + n}" name="{key}"/>'
            f'{graphic(key, cx, cy, n)}</wp:inline></w:drawing></w:r>')


def run(text, bold=False, sz=None, color=None):
    rpr = ('<w:b/>' if bold else '') + (f'<w:color w:val="{color}"/>' if color else '') + (f'<w:sz w:val="{sz}"/>' if sz else '')
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(inner, before=0, after=60, keep=False, brk=False, jc='both'):
    sp = f'<w:spacing w:before="{before}" w:after="{after}"/>' if before else f'<w:spacing w:after="{after}"/>'
    return (f'<w:p><w:pPr>{"<w:keepLines/>" if keep else ""}{"<w:pageBreakBefore/>" if brk else ""}{sp}'
            f'<w:jc w:val="{jc}"/></w:pPr>{inner}</w:p>')


def table(rows, widths, bold_rows=(0,), shade_rows=(0,), total_rows=(), indent=0):
    border = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
                     for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    xml = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>'
           + (f'<w:tblInd w:w="{indent}" w:type="dxa"/>' if indent else '')
           + f'<w:tblBorders>{border}</w:tblBorders><w:tblLayout w:type="fixed"/>'
           '<w:tblCellMar><w:top w:w="8" w:type="dxa"/><w:left w:w="40" w:type="dxa"/><w:bottom w:w="8" w:type="dxa"/>'
           '<w:right w:w="40" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'
           + ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths) + '</w:tblGrid>')
    for i, row in enumerate(rows):
        xml += '<w:tr><w:trPr><w:cantSplit/></w:trPr>'
        for j, (cell, w) in enumerate(zip(row, widths)):
            fill = 'D9D9D9' if i in shade_rows else ('F2F2F2' if i in total_rows else None)
            shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ''
            b = i in bold_rows or i in total_rows or (j == 0 and i not in shade_rows and False)
            jc = '' if j == 0 else '<w:jc w:val="center"/>'
            xml += (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}<w:vAlign w:val="center"/></w:tcPr>'
                    f'<w:p><w:pPr><w:spacing w:after="0"/>{jc}</w:pPr>{run(cell, bold=b, sz=15) if cell else ""}</w:p></w:tc>')
        xml += '</w:tr>'
    return xml + '</w:tbl>'


GAP_BEFORE = 80
body, n = [], 0


def section(label, paras, chart=None, first=False, brk=False):
    global n
    lead = ''
    if chart:
        n += 1
        lead = anchor(chart, n)
    body.append(para(lead + run(label + ' ', True) + run(paras[0]), before=0 if (first or brk) else GAP_BEFORE,
                     keep=bool(chart), brk=brk))
    for p in paras[1:]:
        body.append(para(run(p)))


for i, (label, paras, chart) in enumerate(M.PAGE1):
    section(label, paras, chart, first=(i == 0))
for label, paras, chart in M.PAGE2:
    section(label, paras, chart, brk=True)

# split section: text + KPI table on the left, bridge chart on the right (borderless layout table)
LW, RW = 6250, 4550
n += 1
left = (para(run(M.SPLIT_LABEL + ' ', True) + run(M.SPLIT_TEXT), after=60)
        + table(M.KPI_TABLE, [2700, 1150, 1150, 1000], total_rows=(len(M.KPI_TABLE) - 1,))
        + '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>')
right = para(inline('ex4', 3.12, n), after=0, jc='right')
nob = ''.join(f'<w:{s} w:val="nil"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
body.append(para('', before=0, after=0))
body.append(f'<w:tbl><w:tblPr><w:tblW w:w="{LW + RW}" w:type="dxa"/><w:tblBorders>{nob}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{LW}"/><w:gridCol w:w="{RW}"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr>'
            f'<w:tc><w:tcPr><w:tcW w:w="{LW}" w:type="dxa"/></w:tcPr>{left}</w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="{RW}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>{right}</w:tc></w:tr></w:tbl>')

for label, paras, extra in M.AFTER:
    section(label, paras)
    if extra == 'table':
        body.append(table(M.SCEN_TABLE, [3700, 650, 1150, 1150, 1050, 1000, 1000], total_rows=(len(M.SCEN_TABLE) - 1,)))
        body.append('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>')
body.append('<w:p><w:pPr><w:spacing w:before="60" w:after="0"/><w:jc w:val="both"/></w:pPr>'
            f'<w:r><w:rPr><w:color w:val="404040"/><w:sz w:val="13"/></w:rPr><w:t xml:space="preserve">{escape(M.SOURCES)}</w:t></w:r></w:p>')

x = head + ''.join(body) + sect + '</w:body></w:document>'
open(docp, 'w', encoding='utf-8').write(x)
print('built', out)

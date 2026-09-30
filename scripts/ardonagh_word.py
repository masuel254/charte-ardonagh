#!/usr/bin/env python3
"""Charte Ardonagh (maquette A) pour tout .docx. Le fond n'est jamais modifié.
Usage :
  python3 ardonagh_word.py inspect in.docx
  python3 ardonagh_word.py apply in.docx out.docx params.json
"""
import re, os, sys, json, shutil, subprocess, zipfile, colorsys, tempfile
from xml.sax.saxutils import escape as esc
from lxml import etree

# ------------------------------------------------------------------ charte
GREEN, DEEP, MID, SOFT, SOFT2, LINE, SUB, GOLD, SAND, RED, REDSOFT, BAR, TAUPE = (
    '093428', '14211C', '4B734F', 'EAEFEC', 'F1F5F2', 'DDDCD8', '5F5F5F', 'ECD383', 'D6D2CB', '9C2B1F', 'FAF0EE', 'C8C8C7', 'A39C91')
CHARTER = {GREEN, DEEP, MID, SOFT, SOFT2, LINE, SUB, GOLD, SAND, RED, REDSOFT, BAR, TAUPE, 'FFFFFF', 'D9D9D9', '183028'}
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'w': W_NS, 'r': R_NS}
q = lambda t: '{%s}%s' % (W_NS, t)

def classify(h, kind):
    """Ramène n'importe quelle couleur dans la palette Ardonagh. kind = text | fill | border."""
    h = h.upper()
    if h in CHARTER or not re.fullmatch(r'[0-9A-F]{6}', h): return h
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    H, L, S = colorsys.rgb_to_hls(r, g, b); deg = H * 360
    if S < 0.13 or max(r, g, b) - min(r, g, b) < 0.08:          # gris
        if kind == 'text': return DEEP if L < 0.35 else SUB if L < 0.75 else h
        if kind == 'border': return DEEP if L < 0.35 else LINE
        return DEEP if L < 0.3 else SUB if L < 0.55 else BAR if L < 0.8 else SOFT2 if L < 0.97 else h
    warm, yellow = (deg < 40 or deg >= 330), (40 <= deg < 70)
    if L >= 0.84:                                                # teinte pastel
        if kind == 'text': return h
        if kind == 'border': return LINE
        return REDSOFT if warm else SOFT2 if yellow else SOFT
    if warm: return RED
    if yellow: return GOLD if kind != 'text' else MID
    if kind == 'border' and L > 0.7: return LINE
    if kind == 'fill' and L > 0.7: return SOFT
    return GREEN if L < 0.38 else MID

def remap_xml(x):
    x = re.sub(r'(<w:color [^>]*?w:val=")([0-9A-Fa-f]{6})"', lambda m: m.group(1) + classify(m.group(2), 'text') + '"', x)
    x = re.sub(r'(w:fill=")([0-9A-Fa-f]{6})"', lambda m: m.group(1) + classify(m.group(2), 'fill') + '"', x)
    x = re.sub(r'(<w:(?:top|left|bottom|right|insideH|insideV|between|bar|start|end)\b[^>]*?w:color=")([0-9A-Fa-f]{6})"',
               lambda m: m.group(1) + classify(m.group(2), 'border') + '"', x)
    x = re.sub(r'(<w:shd [^>]*?w:color=")([0-9A-Fa-f]{6})"', lambda m: m.group(1) + classify(m.group(2), 'fill') + '"', x)
    x = re.sub(r'\s+w:theme(?:Color|Shade|Tint|Fill|FillShade|FillTint)="[^"]*"', '', x)   # la couleur explicite prime
    return x

NARROW = r'Calibri|Calibri Light|Aptos|Aptos Display|Aptos Narrow|Carlito|Segoe UI|Segoe UI Light|Cambria|Times New Roman|Georgia|Verdana|Tahoma|Helvetica|Open Sans|Roboto|Lato|Montserrat|Century Gothic|Garamond|Book Antiqua'
def fonts(x):
    x = re.sub(r'(w:(?:ascii|hAnsi|cs|eastAsia)=")(' + NARROW + ')"', r'\1Arial"', x)
    return re.sub(r'\s+w:(?:ascii|hAnsi|eastAsia|cs)Theme="[^"]*"', '', x)

def scale_sz(x, k):
    if abs(k - 1) < 0.01: return x
    return re.sub(r'<w:(sz|szCs) w:val="(\d+)"/>', lambda m: f'<w:{m.group(1)} w:val="{max(12, round(int(m.group(2)) * k))}"/>', x)

# ------------------------------------------------------------------ fabrique OOXML
def run(t, sz=19, b=False, i=False, color='FFFFFF', sp=None):
    rpr = ('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' + ('<w:b/><w:bCs/>' if b else '') + ('<w:i/><w:iCs/>' if i else '')
           + f'<w:color w:val="{color}"/>' + (f'<w:spacing w:val="{sp}"/>' if sp else '') + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r>'

def para(runs, jc='left', before=0, after=0, bdr=''):
    return f'<w:p><w:pPr>{bdr}<w:spacing w:before="{before}" w:after="{after}" w:line="264" w:lineRule="auto"/><w:jc w:val="{jc}"/></w:pPr>{runs}</w:p>'

def cell(w, content, fill=None, mar=(0, 0, 0, 0), borders='', valign='top', span=1):
    t, l, b, r = mar
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>' + (f'<w:gridSpan w:val="{span}"/>' if span > 1 else '') + borders
            + (f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else '')
            + f'<w:tcMar><w:top w:w="{t}" w:type="dxa"/><w:left w:w="{l}" w:type="dxa"/><w:bottom w:w="{b}" w:type="dxa"/><w:right w:w="{r}" w:type="dxa"/></w:tcMar>'
            f'<w:vAlign w:val="{valign}"/></w:tcPr>{content}</w:tc>')

def tbl(widths, rows, ind=0, jc=None):
    nil = ''.join(f'<w:{s} w:val="nil"/>' for s in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'])
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>' + (f'<w:jc w:val="{jc}"/>' if jc else '')
            + f'<w:tblInd w:w="{ind}" w:type="dxa"/><w:tblBorders>{nil}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr>'
            '<w:tblGrid>' + ''.join('<w:gridCol w:w="%d"/>' % w for w in widths) + '</w:tblGrid>' + ''.join(rows) + '</w:tbl>')

def tr(cells, h=None):
    return f'<w:tr>' + (f'<w:trPr><w:trHeight w:val="{h}" w:hRule="exact"/></w:trPr>' if h else '') + f'{cells}</w:tr>'

def drawing(rid, wpx, hpx, did, floating=False):
    cx, cy = int(wpx * 9525), int(hpx * 9525)
    g = (f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
         f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="{did}" name="ard{did}"/><pic:cNvPicPr/></pic:nvPicPr>'
         f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
         f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')
    if floating:
        return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{251658240 + did}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
                f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
                f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="{did}" name="Image {did}"/><wp:cNvGraphicFramePr/>{g}</wp:anchor></w:drawing></w:r>')
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{did}" name="Logo {did}"/><wp:cNvGraphicFramePr/>{g}</wp:inline></w:drawing></w:r>')

WB = '<w:tcBorders>' + ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="FFFFFF"/>' for s in ['top', 'left', 'bottom', 'right']) + '</w:tcBorders>'
TINY = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="14" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'

def cover_xml(c, W, H, has_photo, has_logo):
    """Couverture pleine page (slide 1 du one-pager) : photo à gauche, aplat vert, logo blanc, encadré de suivi."""
    L = 3000 if has_photo else 0
    rule = lambda side: f'<w:pBdr><w:{side} w:val="single" w:sz="8" w:space="10" w:color="FFFFFF"/></w:pBdr>'
    vc = ''
    if c.get('vc'):
        vw = [1500, 2400]
        rows = [tr(cell(sum(vw), para(run(c.get('vc_title', 'Suivi du document'), 14, b=True)), mar=(40, 90, 40, 90), borders=WB, span=2))]
        rows += [tr(cell(vw[0], para(run(k, 13, b=True, color=SAND)), mar=(30, 90, 30, 90), borders=WB) + cell(vw[1], para(run(v, 13)), mar=(30, 90, 30, 90), borders=WB)) for k, v in c['vc']]
        vc = tbl(vw, rows, jc='right')
    logo = (para(drawing('rIdArdLogo', 250, 89, 9002), jc='center', before=1500, after=2300) if has_logo else
            para(run('THE', 16, sp=60), jc='center', before=1700, after=0) + para(run('ARDONAGH', 52, sp=40), jc='center', after=0)
            + para(run('GROUP', 16, sp=60), jc='center', after=2300))
    inner = ((para(drawing('rIdArdPhoto', 200, 1123, 9001, floating=True)) if has_photo else para(''))
             + para('', after=700) + vc + logo
             + para(run(c['strap'], 13, b=True, color=GOLD, sp=16), jc='center', after=160, bdr=rule('top'))
             + para(run(c['kicker'], 26, b=True, sp=20), jc='center', after=60)
             + (para(run(c['title'], 28, i=True), jc='center', after=60) if c.get('title') else '')
             + ''.join(para(run(s, 17, color=SAND), jc='center', after=40) for s in c.get('subs', []))
             + para(run(c.get('date', ''), 15, b=True), jc='right', before=160, after=0, bdr=rule('top')))
    cover = tbl([W], [tr(cell(W, inner, fill=GREEN, mar=(0, L + 900, 0, 900)), h=H - 40)])
    sect = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="14" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr><w:sectPr>'
            f'<w:pgSz w:w="{W}" w:h="{H}"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>'
            '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:pPr></w:p>')
    return cover + sect

HF_NS = ('xmlns:w="%s" xmlns:r="%s" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
         'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' % (W_NS, R_NS))

def header_xml(c, pw, lm, cw):
    band = tbl([pw], [tr(cell(pw, para(''), fill=DEEP), h=180),
                      tr(cell(pw, f'<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="{cw}"/></w:tabs><w:spacing w:before="0" w:after="0"/></w:pPr>'
                              + run(c['band'], c.get('band_sz', 28)) + '<w:r><w:tab/></w:r>' + run(c.get('band_right', ''), 15, b=True, color=GOLD, sp=20) + '</w:p>',
                              fill=GREEN, mar=(0, lm, 0, lm), valign='center'), h=900)], ind=-lm)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {HF_NS}>{band}{TINY}</w:hdr>'

def fld(instr, sz, b, color):
    rp = f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' + ('<w:b/>' if b else '') + f'<w:color w:val="{color}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    return (f'<w:r>{rp}<w:fldChar w:fldCharType="begin"/></w:r><w:r>{rp}<w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            f'<w:r>{rp}<w:fldChar w:fldCharType="separate"/></w:r><w:r>{rp}<w:t>1</w:t></w:r><w:r>{rp}<w:fldChar w:fldCharType="end"/></w:r>')

def footer_xml(c, pw, lm, cw):
    p = (f'<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="{cw}"/></w:tabs><w:spacing w:before="0" w:after="100"/></w:pPr>'
         + run(c['footer'], 14, color=SUB) + '<w:r><w:tab/></w:r>' + fld('PAGE', 16, True, GREEN) + run(' / ', 16, color=SUB) + fld('NUMPAGES', 16, False, SUB) + '</w:p>')
    bar = tbl([pw], [tr(cell(pw, para(''), fill=BAR), h=170)], ind=-lm)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {HF_NS}>{p}{bar}{TINY}</w:ftr>'

# ------------------------------------------------------------------ styles
PPR_AFTER_BDR = ['<w:shd ', '<w:tabs>', '<w:suppressAutoHyphens', '<w:snapToGrid', '<w:spacing ', '<w:ind ', '<w:contextualSpacing', '<w:jc ',
                 '<w:textAlignment', '<w:outlineLvl', '<w:rPr>', '<w:sectPr', '<w:pPrChange', '</w:pPr>']
def add_pbdr(xml, bdr):
    if '<w:pPr>' not in xml:
        xml = re.sub(r'(<w:p(?: [^>]*)?>|<w:style [^>]*>(?:<w:name [^>]*/>))', r'\1<w:pPr></w:pPr>', xml, 1)
    a = xml.find('<w:pPr>'); b = xml.find('</w:pPr>', a); ppr = xml[a:b + 8]
    pos = min(ppr.find(t) for t in PPR_AFTER_BDR if ppr.find(t) > 0)
    return xml[:a] + ppr[:pos] + bdr + ppr[pos:] + xml[b + 8:]

RPR_ORDER = ['<w:rStyle', '<w:rFonts', '<w:b/>', '<w:b ', '<w:bCs', '<w:i/>', '<w:i ', '<w:iCs', '<w:caps', '<w:smallCaps', '<w:strike', '<w:dstrike', '<w:outline',
             '<w:shadow', '<w:emboss', '<w:imprint', '<w:noProof', '<w:snapToGrid', '<w:vanish', '<w:webHidden', '<w:color', '<w:spacing', '<w:w ', '<w:kern',
             '<w:position', '<w:sz ', '<w:szCs', '<w:highlight', '<w:u ', '<w:effect', '<w:bdr', '<w:shd', '<w:fitText', '<w:vertAlign', '<w:rtl', '<w:cs/>',
             '<w:em ', '<w:lang', '<w:eastAsianLayout', '<w:specVanish', '<w:oMath']
def set_rpr(style, bold=None, color=None, sz=None):
    """Réécrit gras / couleur / taille dans le rPr d'un style, en respectant l'ordre du schéma."""
    if '<w:rPr>' not in style: style = style.replace('</w:style>', '<w:rPr></w:rPr></w:style>')
    a = style.rfind('<w:rPr>'); b = style.find('</w:rPr>', a); body = style[a + 7:b]
    items = re.findall(r'<w:[A-Za-z]+(?: [^>]*)?/>|<w:[A-Za-z]+(?: [^>]*)?>.*?</w:[A-Za-z]+>', body, re.S)
    drop = []
    if bold is not None: drop += ['b', 'bCs']
    if color: drop += ['color']
    if sz: drop += ['sz', 'szCs']
    items = [i for i in items if re.match(r'<w:([A-Za-z]+)', i).group(1) not in drop]
    if bold: items += ['<w:b/>', '<w:bCs/>']
    if color: items.append(f'<w:color w:val="{color}"/>')
    if sz: items += [f'<w:sz w:val="{sz}"/>', f'<w:szCs w:val="{sz}"/>']
    rank = lambda i: next((k for k, t in enumerate(RPR_ORDER) if i.startswith(t)), 99)
    items.sort(key=rank)
    return style[:a + 7] + ''.join(items) + style[b:]

def style_ids(styles):
    out = {}
    for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]*)"[^>]*>\s*<w:name w:val="([^"]*)"', styles):
        out[m.group(2).lower()] = m.group(1)
    return out

def restyle_styles(st):
    st = remap_xml(fonts(st))
    st = st.replace('<w:rPrDefault/>', '<w:rPrDefault><w:rPr></w:rPr></w:rPrDefault>')
    if '<w:rPrDefault>' not in st and '<w:docDefaults>' in st:
        st = st.replace('<w:docDefaults>', '<w:docDefaults><w:rPrDefault><w:rPr></w:rPr></w:rPrDefault>', 1)
    elif '<w:docDefaults>' not in st:
        st = re.sub(r'(<w:styles[^>]*>)', r'\1<w:docDefaults><w:rPrDefault><w:rPr></w:rPr></w:rPrDefault></w:docDefaults>', st, 1)
    m = re.search(r'<w:rPrDefault>(.*?)</w:rPrDefault>', st, re.S)
    blk = '<w:style>' + (m.group(1) if '<w:rPr>' in m.group(1) else '<w:rPr></w:rPr>') + '</w:style>'
    blk = re.sub(r'<w:rFonts [^>]*/>', '', blk)
    blk = set_rpr(blk, color=DEEP).replace('<w:rPr>', '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>', 1)
    st = st[:m.start(1)] + blk[len('<w:style>'):-len('</w:style>')] + st[m.end(1):]
    ids = style_ids(st)
    def edit(name, fn):
        nonlocal st
        sid = ids.get(name)
        if not sid: return
        m = re.search(rf'<w:style [^>]*w:styleId="{re.escape(sid)}".*?</w:style>', st, re.S)
        st = st[:m.start()] + fn(m.group(0)) + st[m.end():]
    def h1(s):
        s = re.sub(r'<w:pBdr>.*?</w:pBdr>', '', s, flags=re.S); s = re.sub(r'<w:shd [^>]*/>', '', s)
        s = add_pbdr(s, f'<w:pBdr><w:bottom w:val="single" w:sz="24" w:space="4" w:color="{GREEN}"/></w:pBdr>')
        return set_rpr(s, bold=True, color=DEEP, sz=26)
    def h2(s):
        s = re.sub(r'<w:pBdr>.*?</w:pBdr>', '', s, flags=re.S); s = re.sub(r'<w:shd [^>]*/>', '', s)
        s = add_pbdr(s, f'<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="3" w:color="{LINE}"/></w:pBdr>')
        return set_rpr(s, bold=True, color=GREEN)
    edit('heading 1', h1); edit('heading 2', h2)
    for name in list(ids):
        if re.search(r'toc heading|titre ?sommaire|^sommaire$|titre table des mati', name): edit(name, lambda s: h1(re.sub(r'<w:ind [^>]*/>', '', s)))
    edit('heading 3', lambda s: set_rpr(re.sub(r'<w:shd [^>]*/>', '', s), bold=True, color=MID))
    edit('title', lambda s: set_rpr(s, bold=True, color=GREEN))
    edit('subtitle', lambda s: set_rpr(s, color=SUB))
    edit('hyperlink', lambda s: set_rpr(s, color=MID))
    return st, ids

def bullets(num):
    def lvl(m):
        l = m.group(0)
        if not re.match(r'<w:lvl w:ilvl="0"', l) or 'w:val="bullet"' not in l: return l
        l = re.sub(r'<w:lvlText w:val="[^"]*"/>', '<w:lvlText w:val=""/>', l)
        rpr = f'<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/><w:color w:val="{GREEN}"/><w:sz w:val="14"/><w:szCs w:val="14"/></w:rPr>'
        return re.sub(r'<w:rPr>.*?</w:rPr>', rpr, l, flags=re.S) if '<w:rPr>' in l else l.replace('</w:lvl>', rpr + '</w:lvl>')
    num = re.sub(r'<w:lvl [^>]*>.*?</w:lvl>', lvl, num, flags=re.S)
    num = remap_xml(fonts(num))
    return re.sub(r'(<w:color w:val=")(' + RED + '|' + MID + ')"', rf'\g<1>{GREEN}"', num)   # puces et numéros toujours verts

THEME = {'dk2': '183028', 'lt2': 'FFFFFF', 'accent1': GREEN, 'accent2': 'E5B9A7', 'accent3': BAR, 'accent4': GOLD, 'accent5': TAUPE, 'accent6': 'D6D2CB', 'hlink': MID}
def restyle_theme(th):
    for k, v in THEME.items():
        th = re.sub(rf'(<a:{k}>)(.*?)(</a:{k}>)', rf'\1<a:srgbClr val="{v}"/>\3', th, flags=re.S)
    return re.sub(r'(<a:(?:major|minor)Font><a:latin typeface=")[^"]*"', r'\1Arial"', th)

# ------------------------------------------------------------------ corps
def body_edits(doc_xml, ids, cover_end, plain_tables=True):
    root = etree.fromstring(doc_xml.encode())
    body = root.find('w:body', NS)
    kids = [k for k in body if k.tag != q('sectPr')]
    for k in kids[:cover_end]: body.remove(k)
    # saut de page orphelin en tête de corps
    first = next((k for k in body if k.tag != q('sectPr')), None)
    while first is not None and first.tag == q('p') and not ''.join(first.itertext()).strip() and first.find('.//w:br[@w:type="page"]', NS) is not None and first.find('.//w:sectPr', NS) is None:
        body.remove(first); first = next((k for k in body if k.tag != q('sectPr')), None)
    if first is not None:
        for pb in first.findall('.//w:pageBreakBefore', NS): pb.getparent().remove(pb)
    ph = etree.SubElement(body, q('p')); etree.SubElement(etree.SubElement(ph, q('r')), q('t')).text = '@@COVER@@'
    body.remove(ph); body.insert(0, ph)
    # titres : on retire la mise en forme directe qui contredirait la charte
    hid = {ids.get('heading 1'), ids.get('heading 2'), ids.get('toc heading')} - {None}
    for p in body.iter(q('p')):
        ps = p.find('w:pPr/w:pStyle', NS)
        if ps is not None and ps.get(q('val')) in hid:
            ppr = p.find('w:pPr', NS)
            for t in ('pBdr', 'shd'):
                for e in ppr.findall('w:' + t, NS): ppr.remove(e)
            for rpr in p.findall('.//w:r/w:rPr', NS):
                for t in ('color', 'shd', 'highlight'):
                    for e in rpr.findall('w:' + t, NS): rpr.remove(e)
    # tableaux sans aucune couleur : en-tête vert, filets gris clair
    if plain_tables:
        for t in body.iter(q('tbl')):
            if any(s.get(q('fill'), 'auto').upper() not in ('AUTO', 'FFFFFF') for s in t.iter(q('shd'))): continue
            if 'PAGEREF' in etree.tostring(t).decode() or t.find('.//w:tbl', NS) is not None: continue      # sommaire ou tableau de mise en page
            rows = t.findall('w:tr', NS)
            if len(rows) < 2 or max(len(r.findall('w:tc', NS)) for r in rows) < 2: continue
            tp = t.find('w:tblPr', NS)
            tb = tp.find('w:tblBorders', NS); ts = tp.find('w:tblStyle', NS)
            vis = tb is not None and any(e.get(q('val')) not in ('nil', 'none') for e in tb)
            if not vis and ts is None: continue                                                           # tableau invisible = mise en page
            for e in tp.findall('w:tblBorders', NS): tp.remove(e)
            bd = etree.fromstring(f'<w:tblBorders xmlns:w="{W_NS}">' + ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="{LINE}"/>' for s in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) + '</w:tblBorders>')
            anchor = next((tp.find('w:' + n, NS) for n in ['shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption'] if tp.find('w:' + n, NS) is not None), None)
            anchor.addprevious(bd) if anchor is not None else tp.append(bd)
            for e in tp.findall('w:tblCellMar', NS): tp.remove(e)
            cm = etree.fromstring(f'<w:tblCellMar xmlns:w="{W_NS}"><w:top w:w="50" w:type="dxa"/><w:left w:w="110" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tblCellMar>')
            lk = next((tp.find('w:' + n, NS) for n in ['tblLook', 'tblCaption', 'tblDescription'] if tp.find('w:' + n, NS) is not None), None)
            lk.addprevious(cm) if lk is not None else tp.append(cm)
            row = t.find('w:tr', NS)
            if row is None: continue
            for tc in row.findall('w:tc', NS):
                tcpr = tc.find('w:tcPr', NS)
                if tcpr is None: tcpr = etree.SubElement(tc, q('tcPr')); tc.remove(tcpr); tc.insert(0, tcpr)
                for e in tcpr.findall('w:shd', NS): tcpr.remove(e)
                shd = etree.fromstring(f'<w:shd xmlns:w="{W_NS}" w:val="clear" w:color="auto" w:fill="{GREEN}"/>')
                nxt = next((tcpr.find('w:' + n, NS) for n in ['noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'] if tcpr.find('w:' + n, NS) is not None), None)
                nxt.addprevious(shd) if nxt is not None else tcpr.append(shd)
                for r in tc.iter(q('r')):
                    rpr = r.find('w:rPr', NS)
                    if rpr is None: rpr = etree.Element(q('rPr')); r.insert(0, rpr)
                    for e in rpr.findall('w:color', NS) + rpr.findall('w:b', NS): rpr.remove(e)
                    b = etree.fromstring(f'<w:b xmlns:w="{W_NS}"/>'); c = etree.fromstring(f'<w:color xmlns:w="{W_NS}" w:val="FFFFFF"/>')
                    rf = rpr.find('w:rFonts', NS) if rpr.find('w:rFonts', NS) is not None else rpr.find('w:rStyle', NS)
                    (rf.addnext(b) if rf is not None else rpr.insert(0, b))
                    later = next((rpr.find('w:' + n, NS) for n in ['spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish'] if rpr.find('w:' + n, NS) is not None), None)
                    later.addprevious(c) if later is not None else rpr.append(c)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True).decode()

def restyle_sections(doc):
    """Chaque section reçoit le bandeau et le pied Ardonagh adaptés à sa largeur de page."""
    made = {}
    def fix(m):
        sp = m.group(0)
        if '@@SKIP@@' in sp: return sp
        pw = int(re.search(r'<w:pgSz [^>]*w:w="(\d+)"', sp).group(1))
        mar = re.search(r'<w:pgMar [^>]*/>', sp).group(0)
        lm = int(re.search(r'w:left="(\d+)"', mar).group(1)); rm = int(re.search(r'w:right="(\d+)"', mar).group(1))
        key = (pw, lm, rm); made.setdefault(key, len(made) + 1)
        sp = sp.replace(mar, f'<w:pgMar w:top="1650" w:right="{rm}" w:bottom="1150" w:left="{lm}" w:header="0" w:footer="0" w:gutter="0"/>')
        sp = re.sub(r'<w:headerReference [^>]*/>|<w:footerReference [^>]*/>|<w:titlePg/>|<w:titlePg w:val="[^"]*"/>', '', sp)
        n = made[key]
        return re.sub(r'(<w:sectPr[^>]*>)', rf'\1<w:headerReference w:type="default" r:id="rIdArdHdr{n}"/><w:footerReference w:type="default" r:id="rIdArdFtr{n}"/>', sp, 1)
    doc = re.sub(r'<w:sectPr(?: [^>]*)?>.*?</w:sectPr>', fix, doc, flags=re.S)
    return doc, made

# ------------------------------------------------------------------ images (schémas à aplats uniquement)
def recolor_png(path):
    import numpy as np
    from PIL import Image
    im = Image.open(path)
    if im.format != 'PNG' or min(im.size) < 300: return False
    rgb = im.convert('RGB'); a = np.asarray(rgb).reshape(-1, 3)
    cols, counts = np.unique(a[::7], axis=0, return_counts=True)
    order = np.argsort(-counts); top = cols[order[:14]]; share = counts[order[:14]].sum() / counts.sum()
    if share < 0.85: return False                                    # photo ou dégradé : on ne touche pas
    src = ['%02X%02X%02X' % tuple(c) for c in top]
    dst = [classify(h, 'fill') if sum(int(h[i:i + 2], 16) for i in (0, 2, 4)) > 600 else classify(h, 'text') for h in src]
    if all(s == d for s, d in zip(src, dst)): return False
    S = np.array([[int(h[i:i + 2], 16) for i in (0, 2, 4)] for h in src], float); T = np.array([[int(h[i:i + 2], 16) for i in (0, 2, 4)] for h in dst], float)
    P = a.astype(np.float32); out = P.copy(); best = np.full(len(P), 1e9, np.float32)
    for i in range(len(S)):
        for j in range(i + 1, len(S)):
            d = S[j] - S[i]; dd = (d * d).sum()
            if dd == 0: continue
            for s0 in range(0, len(P), 1_000_000):
                Q = P[s0:s0 + 1_000_000]; t = np.clip(((Q - S[i]) @ d) / dd, 0, 1)
                r = ((Q - (S[i] + t[:, None] * d)) ** 2).sum(1); m = r < best[s0:s0 + 1_000_000]
                best[s0:s0 + 1_000_000][m] = r[m]; out[s0:s0 + 1_000_000][m] = (T[i] + t[:, None] * (T[j] - T[i]))[m]
    out[best > 300] = P[best > 300]
    res = Image.fromarray(out.reshape(np.asarray(rgb).shape).clip(0, 255).astype('uint8'))
    if im.mode == 'RGBA': res.putalpha(im.getchannel('A'))
    res.save(path, optimize=True)
    return True

# ------------------------------------------------------------------ inspection
def inspect(src):
    z = zipfile.ZipFile(src); doc = z.read('word/document.xml').decode(); st = z.read('word/styles.xml').decode()
    root = etree.fromstring(doc.encode()); body = root.find('w:body', NS)
    ids = {v: k for k, v in style_ids(st).items()}
    print('Sections :', len(re.findall(r'<w:sectPr', doc)), '| tableaux :', len(body.findall('.//w:tbl', NS)), '| images :', len([n for n in z.namelist() if n.startswith('word/media/')]))
    print('Polices :', sorted(set(re.findall(r'w:ascii="([^"]*)"', doc + st))))
    print('Couleurs les plus utilisées :', sorted({*re.findall(r'w:(?:val|fill|color)="([0-9A-Fa-f]{6})"', doc)})[:40])
    print('TOC champ :', 'TOC \\' in doc or 'w:instr="TOC' in doc, '| Sommaire statique :', bool(re.search(r'>(Sommaire|SOMMAIRE|Table des matières)<', doc)))
    print('\nPremiers éléments du corps (index : style | texte) :')
    for i, k in enumerate(body):
        if i > 40: break
        if k.tag == q('sectPr'): continue
        ps = k.find('w:pPr/w:pStyle', NS); sname = ids.get(ps.get(q('val')), ps.get(q('val'))) if ps is not None else ''
        brk = ' [SAUT DE PAGE]' if k.find('.//w:br[@w:type="page"]', NS) is not None else ''
        sec = ' [FIN DE SECTION]' if k.find('.//w:sectPr', NS) is not None else ''
        kind = 'TABLEAU' if k.tag == q('tbl') else 'P'
        print(f'{i:>3} {kind} {sname[:18]:<18}| {"".join(k.itertext())[:90]}{brk}{sec}')

# ------------------------------------------------------------------ application
def apply(src, out, c):
    wd = tempfile.mkdtemp(prefix='ard_')
    with zipfile.ZipFile(src) as z:
        for n in z.namelist():
            if '..' in n or n.startswith('/'): continue
            z.extract(n, wd)
    w = lambda p: os.path.join(wd, 'word', p)
    rd = lambda p: open(w(p), encoding='utf-8').read()
    def wr(p, s): open(w(p), 'w', encoding='utf-8').write(s)
    assets = c.get('assets_dir', '')
    has_photo = os.path.exists(os.path.join(assets, 'photo.jpg')); has_logo = os.path.exists(os.path.join(assets, 'logo_white.png'))

    doc = rd('document.xml'); st = rd('styles.xml')
    dd = re.search(r'<w:docDefaults>.*?</w:docDefaults>', st, re.S)
    th = open(w('theme/theme1.xml'), encoding='utf-8').read() if os.path.exists(w('theme/theme1.xml')) else ''
    mf = re.search(r'<a:minorFont><a:latin typeface="([^"]*)"', th)
    narrow = bool(re.search(r'Calibri|Aptos|Carlito|Segoe', (dd.group(0) if dd else '') + (mf.group(1) if mf else '')))
    k = c.get('font_scale', 0.92 if narrow else 1.0)

    st, ids = restyle_styles(st); wr('styles.xml', scale_sz(st, k))
    doc = body_edits(doc, ids, int(c.get('cover_end', 0)), c.get('plain_tables', True))
    pgw = int(re.findall(r'<w:pgSz [^>]*w:w="(\d+)"', doc)[-1]); pgh = int(re.findall(r'<w:pgSz [^>]*w:h="(\d+)"', doc)[-1])
    W, H = (pgw, pgh) if pgw < pgh else (11906, 16838)
    doc = doc.replace('<w:lastRenderedPageBreak/>', '')
    doc = scale_sz(remap_xml(fonts(doc)), k)
    doc, made = restyle_sections(doc)
    # sommaire statique : numéros en vert
    m = re.search(r'>(Sommaire|SOMMAIRE|Table des matières|TABLE DES MATIÈRES)</w:t>', doc)
    if m:
        h1 = ids.get('heading 1', '§§'); end = doc.find(f'<w:pStyle w:val="{h1}"/>', m.end())
        if end > 0: doc = doc[:m.start()] + doc[m.start():end].replace(f'w:val="{RED}"', f'w:val="{GREEN}"') + doc[end:]
    doc = re.sub(r'<w:p(?: [^>]*)?><w:r><w:t>@@COVER@@</w:t></w:r></w:p>', lambda _: cover_xml(c['cover'], W, H, has_photo, has_logo), doc)
    wr('document.xml', doc)

    if os.path.exists(w('numbering.xml')): wr('numbering.xml', bullets(rd('numbering.xml')))
    if os.path.exists(w('theme/theme1.xml')): wr('theme/theme1.xml', restyle_theme(rd('theme/theme1.xml')))
    if os.path.exists(w('settings.xml')):
        s = re.sub(r'<w:evenAndOddHeaders(?: [^>]*)?/>', '', rd('settings.xml'))
        if ('TOC \\' in doc or 'TOC \\o' in doc) and '<w:updateFields' not in s:
            s = re.sub(r'(<w:settings[^>]*>)', r'\1<w:updateFields w:val="true"/>', s, 1)
        wr('settings.xml', s)
    # anciens en-têtes / pieds (conservés pour d'éventuelles sections non traitées) : couleurs et polices seulement
    for f in os.listdir(os.path.join(wd, 'word')):
        if re.match(r'(header|footer)\d+\.xml$', f): wr(f, remap_xml(fonts(rd(f))))

    rels = rd('_rels/document.xml.rels'); R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
    ct = open(os.path.join(wd, '[Content_Types].xml'), encoding='utf-8').read(); T = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
    add = ''
    for (pw, lm, rm), n in made.items():
        wr(f'header_ard{n}.xml', header_xml(c, pw, lm, pw - lm - rm)); wr(f'footer_ard{n}.xml', footer_xml(c, pw, lm, pw - lm - rm))
        add += f'<Relationship Id="rIdArdHdr{n}" Type="{R}header" Target="header_ard{n}.xml"/><Relationship Id="rIdArdFtr{n}" Type="{R}footer" Target="footer_ard{n}.xml"/>'
        ct = ct.replace('</Types>', f'<Override PartName="/word/header_ard{n}.xml" ContentType="{T}header+xml"/><Override PartName="/word/footer_ard{n}.xml" ContentType="{T}footer+xml"/></Types>')
    os.makedirs(w('media'), exist_ok=True)
    if has_photo: shutil.copy(os.path.join(assets, 'photo.jpg'), w('media/ard_photo.jpg')); add += f'<Relationship Id="rIdArdPhoto" Type="{R}image" Target="media/ard_photo.jpg"/>'
    if has_logo: shutil.copy(os.path.join(assets, 'logo_white.png'), w('media/ard_logo_white.png')); add += f'<Relationship Id="rIdArdLogo" Type="{R}image" Target="media/ard_logo_white.png"/>'
    wr('_rels/document.xml.rels', rels.replace('</Relationships>', add + '</Relationships>'))
    for ext, mime in (('jpg', 'image/jpeg'), ('png', 'image/png')):
        if f'Extension="{ext}"' not in ct: ct = ct.replace('<Default ', f'<Default Extension="{ext}" ContentType="{mime}"/><Default ', 1)
    open(os.path.join(wd, '[Content_Types].xml'), 'w', encoding='utf-8').write(ct)

    if c.get('recolor_images', True):
        for f in os.listdir(w('media')):
            if f.startswith('ard_') or not f.lower().endswith('.png') or f in c.get('keep_images', []): continue
            if recolor_png(w('media/' + f)): print('Schéma recoloré :', f)

    if os.path.exists(out): os.remove(out)
    subprocess.run(f'cd "{wd}" && zip -q -X -r "{os.path.abspath(out)}" .', shell=True, check=True)
    shutil.rmtree(wd)
    print('OK ->', out)

def body_text(path):
    d = zipfile.ZipFile(path).read('word/document.xml').decode()
    return ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', d))

if __name__ == '__main__':
    if sys.argv[1] == 'inspect': inspect(sys.argv[2])
    elif sys.argv[1] == 'apply': apply(sys.argv[2], sys.argv[3], json.load(open(sys.argv[4], encoding='utf-8')))
    elif sys.argv[1] == 'text': print(body_text(sys.argv[2]))

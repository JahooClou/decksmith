# -*- coding: utf-8 -*-
"""svgshapes: SVG -> native PowerPoint shapes (DrawingML custGeom freeforms), for icons, arrows, diagram parts and
any other vector element that is not a logo. Native shapes recolour in PowerPoint (one fill and one line colour per
shape), scale cleanly and animate like any shape. Logos stay whole: insert them as SVG pictures (svg_logos.ps1).

    from svgshapes import add_svg_shape, trim_svg
    add_svg_shape(slide, open("wifi.svg").read(), Inches(1), Inches(1), Inches(0.5), Inches(0.5),
                  name="wifi icon", colour="C00000")     # one freeform, recolourable
    open("logo_trim.svg", "w").write(trim_svg(open("logo.svg").read()))   # viewBox cropped to the artwork

Reads path, rect, circle, ellipse, line, polyline, polygon; fill / stroke from attributes, inline style, simple CSS
classes and inheritance; transforms (matrix, translate, scale, rotate, skew); arcs and quadratic curves become
cubic Beziers, so geometry stays exact. Single-colour art -> one <p:sp>; several colours -> a <p:grpSp>, one shape
per colour run, in the SVG's stacking order. Not handled: clipPath, mask, gradients, text, embedded images (convert
text to outlines first; flatten clips in the editor).
PowerPoint itself cannot do this through COM: Ungroup on an inserted SVG (msoGraphic) is refused, and Convert to
Shape exists only in the UI.
"""
import math, re
import xml.etree.ElementTree as ET

NS = '{http://www.w3.org/2000/svg}'
NAMED = {'white': 'FFFFFF', 'black': '000000', 'red': 'FF0000', 'none': None}


# ---------------------------------------------------------------- colours and styles
def colour(v, current='000000'):
    if v is None:
        return None
    v = v.strip()
    if v in ('none', 'transparent'):
        return None
    if v == 'currentColor':
        return current
    if v.startswith('#'):
        h = v[1:]
        if len(h) == 3:
            h = ''.join(c * 2 for c in h)
        return h[:6].upper()
    m = re.match(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', v)
    if m:
        return '%02X%02X%02X' % tuple(int(g) for g in m.groups())
    return NAMED.get(v.lower(), '000000')


def css_rules(root):
    rules = {}
    for st in root.iter(NS + 'style'):
        txt = re.sub(r'/\*.*?\*/', '', st.text or '', flags=re.S)
        for sel, body in re.findall(r'([^{}]+)\{([^}]*)\}', txt):
            decl = dict((k.strip(), v.strip()) for k, v in (d.split(':', 1) for d in body.split(';') if ':' in d))
            for s in sel.split(','):
                s = s.strip()
                if s.startswith('.'):
                    rules.setdefault(s[1:], {}).update(decl)
    return rules


PROPS = ('fill', 'stroke', 'stroke-width', 'stroke-linecap', 'stroke-linejoin', 'fill-rule', 'opacity', 'fill-opacity',
         'stroke-opacity', 'color', 'display', 'visibility')


def own_style(el, rules):
    st = {}
    for c in (el.get('class') or '').split():
        st.update({k: v for k, v in rules.get(c, {}).items() if k in PROPS})
    st.update({k: el.get(k) for k in PROPS if el.get(k) is not None})
    if el.get('style'):
        for d in el.get('style').split(';'):
            if ':' in d:
                k, v = d.split(':', 1)
                if k.strip() in PROPS:
                    st[k.strip()] = v.strip()
    return st


# ---------------------------------------------------------------- transforms
def mat_mul(a, b):
    return (a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
            a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5])


def parse_transform(t):
    m = (1, 0, 0, 1, 0, 0)
    for name, args in re.findall(r'(\w+)\s*\(([^)]*)\)', t or ''):
        v = [float(x) for x in re.findall(r'[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?', args)]
        if name == 'matrix':
            n = tuple(v)
        elif name == 'translate':
            n = (1, 0, 0, 1, v[0], v[1] if len(v) > 1 else 0)
        elif name == 'scale':
            n = (v[0], 0, 0, v[1] if len(v) > 1 else v[0], 0, 0)
        elif name == 'rotate':
            a = math.radians(v[0]); c, s = math.cos(a), math.sin(a)
            n = (c, s, -s, c, 0, 0)
            if len(v) == 3:
                n = mat_mul(mat_mul((1, 0, 0, 1, v[1], v[2]), n), (1, 0, 0, 1, -v[1], -v[2]))
        elif name == 'skewX':
            n = (1, 0, math.tan(math.radians(v[0])), 1, 0, 0)
        elif name == 'skewY':
            n = (1, math.tan(math.radians(v[0])), 0, 1, 0, 0)
        else:
            continue
        m = mat_mul(m, n)
    return m


def apply(m, x, y):
    return m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]


# ---------------------------------------------------------------- path data
NUM = re.compile(r'[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?')


def tokens(d):
    for m in re.finditer(r'([MmLlHhVvCcSsQqTtAaZz])|([-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?)', d):
        yield m.group(1) or float(m.group(2))


def arc_to_cubic(x1, y1, rx, ry, phi, fa, fs, x2, y2):
    """Endpoint arc -> list of cubic segments (SVG implementation notes F.6)."""
    if rx == 0 or ry == 0:
        return [(x1, y1, x2, y2, x2, y2)]
    rx, ry = abs(rx), abs(ry)
    cp, sp = math.cos(math.radians(phi)), math.sin(math.radians(phi))
    dx, dy = (x1 - x2) / 2, (y1 - y2) / 2
    x1p, y1p = cp * dx + sp * dy, -sp * dx + cp * dy
    lam = x1p ** 2 / rx ** 2 + y1p ** 2 / ry ** 2
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx ** 2 * ry ** 2 - rx ** 2 * y1p ** 2 - ry ** 2 * x1p ** 2
    den = rx ** 2 * y1p ** 2 + ry ** 2 * x1p ** 2
    co = math.sqrt(max(0, num / den)) if den else 0
    if fa == fs:
        co = -co
    cxp, cyp = co * rx * y1p / ry, -co * ry * x1p / rx
    cx, cy = cp * cxp - sp * cyp + (x1 + x2) / 2, sp * cxp + cp * cyp + (y1 + y2) / 2

    def ang(ux, uy, vx, vy):
        a = math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
        return a
    t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not fs and dt > 0:
        dt -= 2 * math.pi
    elif fs and dt < 0:
        dt += 2 * math.pi
    n = max(1, int(math.ceil(abs(dt) / (math.pi / 2) - 1e-9)))
    d = dt / n
    k = 4 / 3 * math.tan(d / 4)
    out = []
    for i in range(n):
        a1, a2 = t1 + i * d, t1 + (i + 1) * d
        e1 = (math.cos(a1) - k * math.sin(a1), math.sin(a1) + k * math.cos(a1))
        e2 = (math.cos(a2) + k * math.sin(a2), math.sin(a2) - k * math.cos(a2))
        e3 = (math.cos(a2), math.sin(a2))
        pts = []
        for ex, ey in (e1, e2, e3):
            px, py = rx * ex, ry * ey
            pts += [cp * px - sp * py + cx, sp * px + cp * py + cy]
        out.append(tuple(pts))
    return out


def parse_d(d):
    subs, cur = [], None
    x = y = sx = sy = 0.0
    lc = None          # last control point for S / T
    lq = None
    cmd = None
    tk = list(tokens(d))
    i = 0

    def take(n):
        nonlocal i
        v = tk[i:i + n]; i += n
        return v
    while i < len(tk):
        if isinstance(tk[i], str):
            cmd = tk[i]; i += 1
        elif cmd is None:
            break
        rel = cmd.islower(); C = cmd.upper()
        if C == 'Z':
            if cur:
                cur.append(('Z',)); subs.append(cur); cur = None
            x, y = sx, sy; lc = lq = None
            continue
        if C == 'M':
            nx, ny = take(2)
            if rel:
                nx += x; ny += y
            if cur:
                subs.append(cur)
            cur = [('M', nx, ny)]; x, y = sx, sy = nx, ny
            cmd = 'l' if rel else 'L'; lc = lq = None
            continue
        if cur is None:
            cur = [('M', x, y)]; sx, sy = x, y
        if C == 'L':
            nx, ny = take(2)
            if rel:
                nx += x; ny += y
            cur.append(('L', nx, ny)); x, y = nx, ny; lc = lq = None
        elif C == 'H':
            nx = take(1)[0] + (x if rel else 0)
            cur.append(('L', nx, y)); x = nx; lc = lq = None
        elif C == 'V':
            ny = take(1)[0] + (y if rel else 0)
            cur.append(('L', x, ny)); y = ny; lc = lq = None
        elif C == 'C':
            v = take(6)
            if rel:
                v = [v[k] + (x if k % 2 == 0 else y) for k in range(6)]
            cur.append(('C',) + tuple(v)); lc = (v[2], v[3]); x, y = v[4], v[5]; lq = None
        elif C == 'S':
            v = take(4)
            if rel:
                v = [v[k] + (x if k % 2 == 0 else y) for k in range(4)]
            c1 = (2 * x - lc[0], 2 * y - lc[1]) if lc else (x, y)
            cur.append(('C', c1[0], c1[1], v[0], v[1], v[2], v[3])); lc = (v[0], v[1]); x, y = v[2], v[3]; lq = None
        elif C in ('Q', 'T'):
            if C == 'Q':
                v = take(4)
                if rel:
                    v = [v[k] + (x if k % 2 == 0 else y) for k in range(4)]
                q = (v[0], v[1]); ex, ey = v[2], v[3]
            else:
                v = take(2)
                if rel:
                    v = [v[0] + x, v[1] + y]
                q = (2 * x - lq[0], 2 * y - lq[1]) if lq else (x, y); ex, ey = v
            cur.append(('C', x + 2 / 3 * (q[0] - x), y + 2 / 3 * (q[1] - y), ex + 2 / 3 * (q[0] - ex), ey + 2 / 3 * (q[1] - ey), ex, ey))
            lq = q; lc = None; x, y = ex, ey
        elif C == 'A':
            rx, ry, phi, fa, fs, nx, ny = take(7)
            if rel:
                nx += x; ny += y
            for seg in arc_to_cubic(x, y, rx, ry, phi, int(fa), int(fs), nx, ny):
                cur.append(('C',) + seg)
            x, y = nx, ny; lc = lq = None
    if cur:
        subs.append(cur)
    return subs


def ellipse_subpath(cx, cy, rx, ry):
    k = 0.5522847498
    return [[('M', cx + rx, cy),
             ('C', cx + rx, cy + k * ry, cx + k * rx, cy + ry, cx, cy + ry),
             ('C', cx - k * rx, cy + ry, cx - rx, cy + k * ry, cx - rx, cy),
             ('C', cx - rx, cy - k * ry, cx - k * rx, cy - ry, cx, cy - ry),
             ('C', cx + k * rx, cy - ry, cx + rx, cy - k * ry, cx + rx, cy), ('Z',)]]


def f(el, k, d=0.0):
    v = el.get(k)
    return float(NUM.match(v).group(0)) if v else d


def element_subpaths(el, tag):
    if tag == 'path':
        return parse_d(el.get('d') or '')
    if tag == 'rect':
        x, y, w, h = f(el, 'x'), f(el, 'y'), f(el, 'width'), f(el, 'height')
        rx = f(el, 'rx', None) if el.get('rx') else (f(el, 'ry') if el.get('ry') else 0)
        ry = f(el, 'ry', None) if el.get('ry') else rx
        if not rx:
            return [[('M', x, y), ('L', x + w, y), ('L', x + w, y + h), ('L', x, y + h), ('Z',)]]
        rx, ry = min(rx, w / 2), min(ry, h / 2); k = 0.5522847498
        return [[('M', x + rx, y), ('L', x + w - rx, y), ('C', x + w - rx + k * rx, y, x + w, y + ry - k * ry, x + w, y + ry),
                 ('L', x + w, y + h - ry), ('C', x + w, y + h - ry + k * ry, x + w - rx + k * rx, y + h, x + w - rx, y + h),
                 ('L', x + rx, y + h), ('C', x + rx - k * rx, y + h, x, y + h - ry + k * ry, x, y + h - ry),
                 ('L', x, y + ry), ('C', x, y + ry - k * ry, x + rx - k * rx, y, x + rx, y), ('Z',)]]
    if tag == 'circle':
        r = f(el, 'r')
        return ellipse_subpath(f(el, 'cx'), f(el, 'cy'), r, r)
    if tag == 'ellipse':
        return ellipse_subpath(f(el, 'cx'), f(el, 'cy'), f(el, 'rx'), f(el, 'ry'))
    if tag == 'line':
        return [[('M', f(el, 'x1'), f(el, 'y1')), ('L', f(el, 'x2'), f(el, 'y2'))]]
    if tag in ('polyline', 'polygon'):
        v = [float(n) for n in NUM.findall(el.get('points') or '')]
        pts = list(zip(v[0::2], v[1::2]))
        if not pts:
            return []
        sp = [('M',) + pts[0]] + [('L',) + p for p in pts[1:]]
        if tag == 'polygon':
            sp.append(('Z',))
        return [sp]
    return []


def parse(svg_text):
    root = ET.fromstring(svg_text)
    rules = css_rules(root)
    vb = [float(v) for v in NUM.findall(root.get('viewBox') or '')] or [0, 0, f(root, 'width'), f(root, 'height')]
    items = []

    def walk(el, inherited, m):
        tag = el.tag.replace(NS, '')
        if tag in ('defs', 'style', 'title', 'desc', 'clipPath', 'mask', 'symbol', 'metadata'):
            return
        st = dict(inherited); st.update(own_style(el, rules))
        if st.get('display') == 'none' or st.get('visibility') == 'hidden':
            return
        m = mat_mul(m, parse_transform(el.get('transform')))
        if tag in ('path', 'rect', 'circle', 'ellipse', 'line', 'polyline', 'polygon'):
            cur = colour(st.get('color', '#000000'))
            fill = colour(st.get('fill', '#000000'), cur)
            stroke = colour(st.get('stroke', 'none'), cur)
            if tag in ('line', 'polyline'):
                fill = None
            subs = element_subpaths(el, tag)
            if subs and (fill or stroke):
                scale = math.sqrt(abs(m[0] * m[3] - m[1] * m[2]))
                tsubs = []
                for sp in subs:
                    out = []
                    for s in sp:
                        if s[0] == 'Z':
                            out.append(s)
                        else:
                            p = []
                            for k in range(1, len(s), 2):
                                p += apply(m, s[k], s[k + 1])
                            out.append((s[0],) + tuple(p))
                    tsubs.append(out)
                op = float(st.get('opacity', 1))
                items.append({'subpaths': tsubs, 'fill': fill, 'stroke': stroke,
                              'sw': float(NUM.match(st.get('stroke-width', '1')).group(0)) * scale if stroke else 0,
                              'cap': st.get('stroke-linecap', 'butt'), 'join': st.get('stroke-linejoin', 'miter'),
                              'falpha': op * float(st.get('fill-opacity', 1)), 'salpha': op * float(st.get('stroke-opacity', 1))})
        for ch in el:
            walk(ch, st, m)
    walk(root, {}, (1, 0, 0, 1, 0, 0))
    return items, vb


def bbox(items, with_stroke=True):
    xs, ys = [], []
    for it in items:
        h = it['sw'] / 2 if with_stroke else 0
        for sp in it['subpaths']:
            for s in sp:
                for k in range(1, len(s), 2):
                    xs += [s[k] - h, s[k] + h]; ys += [s[k + 1] - h, s[k + 1] + h]
    return min(xs), min(ys), max(xs), max(ys)


# ---------------------------------------------------------------- DrawingML
CAP = {'round': 'rnd', 'square': 'sq', 'butt': 'flat'}
JOIN = {'round': '<a:round/>', 'bevel': '<a:bevel/>', 'miter': '<a:miter lim="800000"/>'}


def _path_xml(subs, x0, y0, sx, sy, W, H, fill, stroke):
    q = lambda vx, vy: (int(round((vx - x0) * sx)), int(round((vy - y0) * sy)))
    out = ['<a:path w="%d" h="%d"%s%s>' % (W, H, '' if fill else ' fill="none"', '' if stroke else ' stroke="0"')]
    for sp in subs:
        for s in sp:
            if s[0] == 'M':
                out.append('<a:moveTo><a:pt x="%d" y="%d"/></a:moveTo>' % q(s[1], s[2]))
            elif s[0] == 'L':
                out.append('<a:lnTo><a:pt x="%d" y="%d"/></a:lnTo>' % q(s[1], s[2]))
            elif s[0] == 'C':
                out.append('<a:cubicBezTo>' + ''.join('<a:pt x="%d" y="%d"/>' % q(s[k], s[k + 1]) for k in (1, 3, 5)) + '</a:cubicBezTo>')
            else:
                out.append('<a:close/>')
    out.append('</a:path>')
    return ''.join(out)


def _alpha(a):
    return '' if a >= 0.999 else '<a:alpha val="%d"/>' % int(round(a * 100000))


def sp_xml(sid, name, items, src_box, emu_box, descr=''):
    """One <p:sp>: items share one fill colour and one line (colour, width, caps)."""
    x0, y0, x1, y1 = src_box
    ox, oy, cx, cy = emu_box
    W, H = 100000, int(round(100000 * (y1 - y0) / (x1 - x0)))
    sx, sy = W / (x1 - x0), H / (y1 - y0)
    fill = next((i['fill'] for i in items if i['fill']), None)
    stroke = next((i for i in items if i['stroke']), None)
    paths = ''.join(_path_xml(i['subpaths'], x0, y0, sx, sy, W, H, i['fill'], i['stroke']) for i in items)
    fx = ('<a:solidFill><a:srgbClr val="%s">%s</a:srgbClr></a:solidFill>' % (fill, _alpha(items[0]['falpha']))) if fill else '<a:noFill/>'
    if stroke:
        w = int(round(stroke['sw'] * cx / (x1 - x0)))
        lx = '<a:ln w="%d" cap="%s"><a:solidFill><a:srgbClr val="%s">%s</a:srgbClr></a:solidFill>%s</a:ln>' % (
            w, CAP.get(stroke['cap'], 'flat'), stroke['stroke'], _alpha(stroke['salpha']), JOIN.get(stroke['join'], JOIN['miter']))
    else:
        lx = '<a:ln><a:noFill/></a:ln>'
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"%s/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>%s</a:pathLst></a:custGeom>'
            '%s%s</p:spPr></p:sp>') % (sid, esc(name), (' descr="%s"' % esc(descr)) if descr else '', ox, oy, cx, cy, paths, fx, lx)


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def style_key(i):
    return (i['fill'], i['stroke'], round(i['sw'], 4) if i['stroke'] else 0, i['cap'], i['join'], round(i['falpha'], 3), round(i['salpha'], 3))


def shape_xml(sid, name, items, src_box, emu_box, next_id, descr=''):
    """Single-colour artwork -> one <p:sp> with id sid. Several colours -> <p:grpSp> with id sid holding one shape
    per colour (ids from next_id()). Shapes keep the SVG's stacking order (runs of the same colour are merged)."""
    runs = []
    for it in items:
        if runs and style_key(runs[-1][0]) == style_key(it):
            runs[-1].append(it)
        else:
            runs.append([it])
    keys = {style_key(r[0]) for r in runs}
    if len(keys) == 1:
        return sp_xml(sid, name, items, src_box, emu_box, descr)
    ox, oy, cx, cy = emu_box
    x0, y0, x1, y1 = src_box
    kx, ky = cx / (x1 - x0), cy / (y1 - y0)
    kids = []
    for r in runs:
        b = bbox(r, with_stroke=False)
        if b[2] - b[0] < 1e-6 or b[3] - b[1] < 1e-6:
            b = (b[0], b[1], max(b[2], b[0] + 1e-3), max(b[3], b[1] + 1e-3))
        eb = (int(round(ox + (b[0] - x0) * kx)), int(round(oy + (b[1] - y0) * ky)), max(1, int(round((b[2] - b[0]) * kx))), max(1, int(round((b[3] - b[1]) * ky))))
        kids.append(sp_xml(next_id(), '%s %d' % (name, len(kids) + 1), r, b, eb).replace(
            '<a:ln w="%d"' % 0, '<a:ln w="0"'))
    # stroke widths inside groups are scaled by kx (sp_xml used the child box), which is the same factor
    return ('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="%d" name="%s"%s/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/><a:chOff x="%d" y="%d"/><a:chExt cx="%d" cy="%d"/></a:xfrm></p:grpSpPr>'
            '%s</p:grpSp>') % (sid, esc(name), (' descr="%s"' % esc(descr)) if descr else '', ox, oy, cx, cy, ox, oy, cx, cy, ''.join(kids))


# ---------------------------------------------------------------- python-pptx helpers
def trim_svg(svg_text, with_stroke=True):
    """The same SVG with its viewBox (and width/height) cropped to the drawn artwork, so a box placed on the slide
    is the artwork's box. Use it for logos before inserting them as SVG pictures."""
    items, vb = parse(svg_text)
    b = bbox(items, with_stroke)
    vbs = '%.3f %.3f %.3f %.3f' % (b[0], b[1], b[2] - b[0], b[3] - b[1])
    t = re.sub(r'viewBox="[^"]*"', 'viewBox="%s"' % vbs, svg_text, count=1)
    if 'viewBox=' not in t:
        t = t.replace('<svg ', '<svg viewBox="%s" ' % vbs, 1)
    for _ in range(2):
        t = re.sub(r'(<svg[^>]*?)\s(width|height)="[^"]*"', r'', t, count=1)
    return t.replace('<svg ', '<svg width="%.0f" height="%.0f" ' % (b[2] - b[0], b[3] - b[1]), 1)


def add_svg_shape(slide, svg_text, x, y, w, h, *, name, colour=None, box='viewbox', descr=''):
    """Append the SVG as native shape(s) to a python-pptx slide; x, y, w, h in EMU (Inches() etc.).
    colour: 'RRGGBB' replaces currentColor and every single colour of a one-colour icon (icon libraries draw with
    currentColor). box='viewbox' maps the whole viewBox to the box (icons: same padding as the library);
    box='content' maps the drawn artwork. Returns the new shape (python-pptx object)."""
    from pptx.oxml import parse_xml
    if colour:
        svg_text = re.sub(r'currentColor', '#' + colour, svg_text)
    items, vb = parse(svg_text)
    if colour and len({(i['fill'], i['stroke']) for i in items}) >= 1:
        for i in items:
            i['fill'] = colour if i['fill'] else None
            i['stroke'] = colour if i['stroke'] else None
    src = (vb[0], vb[1], vb[0] + vb[2], vb[1] + vb[3]) if box == 'viewbox' else bbox(items)
    tree = slide.shapes._spTree
    ids = [int(v) for v in tree.xpath('//p:cNvPr/@id')] or [1]
    nxt = [max(ids) + 1]

    def next_id():
        nxt[0] += 1
        return nxt[0] - 1
    sid = next_id()
    xml = shape_xml(sid, name, items, src, (int(x), int(y), int(w), int(h)), next_id, descr)
    ns = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')
    xml = re.sub(r'^<p:(sp|grpSp)>', lambda m: '<p:%s %s>' % (m.group(1), ns), xml)
    el = parse_xml(xml)
    tree.append(el)
    return next(s for s in slide.shapes if s.shape_id == sid)

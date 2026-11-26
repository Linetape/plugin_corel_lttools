from PIL import Image, ImageDraw, ImageFont
import math, os

OUT = "/home/claude/pacote/LTTools/icones"
os.makedirs(OUT, exist_ok=True)
S = 128
INK = (45, 45, 55, 255)
BEI = (243, 224, 150, 255)
BEI2 = (220, 190, 110, 255)
PUR = (110, 80, 190, 255)
RED = (220, 50, 50, 255)
BLU = (50, 110, 210, 255)
MAG = (225, 40, 180, 255)
YEL = (255, 205, 30, 255)
GRN = (40, 170, 90, 255)
GRY = (150, 150, 160, 255)
W = 7
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"


def font(sz):
    return ImageFont.truetype(BOLD, sz)


def new():
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    return im, ImageDraw.Draw(im)


def rect(d, b, fill=BEI, out=INK, w=W, r=0):
    if r:
        d.rounded_rectangle(b, r, fill=fill, outline=out, width=w)
    else:
        d.rectangle(b, fill=fill, outline=out, width=w)


def circ(d, cx, cy, r, fill=BEI, out=INK, w=W):
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill, outline=out, width=w)


def line(d, pts, c=INK, w=W):
    d.line(pts, fill=c, width=w, joint="curve")


def dashed(d, pts, c=INK, w=6, seg=12):
    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        L = math.hypot(x2 - x1, y2 - y1)
        n = max(1, int(L / seg))
        for k in range(0, n, 2):
            a, b = k / n, min(1, (k + 1) / n)
            d.line([(x1 + (x2 - x1) * a, y1 + (y2 - y1) * a), (x1 + (x2 - x1) * b, y1 + (y2 - y1) * b)], fill=c, width=w)


def text(d, s, cx, cy, sz, c=INK):
    f = font(sz)
    b = d.textbbox((0, 0), s, font=f)
    d.text((cx - (b[2] - b[0]) / 2 - b[0], cy - (b[3] - b[1]) / 2 - b[1]), s, font=f, fill=c)


def arrow(d, x1, y1, x2, y2, c=INK, w=W, h=16):
    d.line([(x1, y1), (x2, y2)], fill=c, width=w)
    a = math.atan2(y2 - y1, x2 - x1)
    for s in (2.6, -2.6):
        d.line([(x2, y2), (x2 + h * math.cos(a + s), y2 + h * math.sin(a + s))], fill=c, width=w)


def curve(d, p0, p1, p2, p3, c=INK, w=W, n=40):
    pts = []
    for i in range(n + 1):
        t = i / n; u = 1 - t
        pts.append((u**3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t**3 * p3[0],
                    u**3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t**3 * p3[1]))
    d.line(pts, fill=c, width=w, joint="curve")
    return pts


def node(d, x, y, c=PUR, r=8):
    d.rectangle([x - r, y - r, x + r, y + r], fill=c, outline=INK, width=3)


I = {}


def icon(name):
    def deco(f):
        I[name] = f
        return f
    return deco

# ---------------- categorias / logo ----------------
@icon("logo")
def _(d):
    rect(d, [6, 14, 122, 114], fill=PUR, out=(70, 40, 140, 255), r=20)
    text(d, "LT", 64, 64, 64, (255, 255, 255, 255))

@icon("cat_medicao")
def _(d):
    rect(d, [10, 70, 118, 112])
    for i, x in enumerate(range(24, 112, 14)):
        line(d, [(x, 70), (x, 84 if i % 2 else 92)], w=5)
    curve(d, (14, 52), (40, 10), (80, 70), (114, 22), c=PUR, w=8)

@icon("cat_encaixe")
def _(d):
    rect(d, [10, 10, 70, 58]); rect(d, [76, 10, 118, 82]); rect(d, [10, 64, 70, 118])
    circ(d, 97, 101, 17, fill=BEI2)

@icon("cat_corte")
def _(d):
    pts = curve(d, (14, 112), (20, 30), (100, 100), (114, 16), c=INK, w=7)
    node(d, 14, 112); node(d, 64, 66); node(d, 114, 16)
    dashed(d, [(14, 112), (30, 40)], c=GRY, w=4); dashed(d, [(114, 16), (100, 90)], c=GRY, w=4)

@icon("cat_letracaixa")
def _(d):
    d.polygon([(20, 18), (52, 18), (52, 86), (100, 86), (100, 114), (20, 114)], fill=BEI2)
    d.polygon([(32, 8), (64, 8), (64, 76), (112, 76), (112, 104), (32, 104)], fill=BEI, outline=INK)
    d.line([(32, 8), (64, 8), (64, 76), (112, 76), (112, 104), (32, 104), (32, 8)], fill=INK, width=W)
    for x, y in [(48, 30), (48, 58), (48, 88), (80, 90), (100, 90)]:
        circ(d, x, y, 6, fill=YEL, out=INK, w=3)

@icon("cat_geradores")
def _(d):
    d.polygon([(64, 14), (114, 38), (64, 62), (14, 38)], fill=BEI, outline=INK)
    d.polygon([(14, 38), (64, 62), (64, 116), (14, 92)], fill=BEI2, outline=INK)
    d.polygon([(64, 62), (114, 38), (114, 92), (64, 116)], fill=(200, 170, 95, 255), outline=INK)
    for p in [[(64, 14), (114, 38), (64, 62), (14, 38), (64, 14)], [(14, 38), (14, 92), (64, 116), (114, 92), (114, 38)], [(64, 62), (64, 116)]]:
        line(d, p, w=6)

@icon("cat_edicao")
def _(d):
    rect(d, [22, 22, 106, 106], fill=BEI)
    for x, y in [(22, 22), (106, 22), (22, 106), (106, 106), (64, 22), (64, 106), (22, 64), (106, 64)]:
        node(d, x, y, c=(255, 255, 255, 255), r=9)
    arrow(d, 64, 64, 118, 118, c=PUR, w=7)

# ---------------- medição ----------------
@icon("area")
def _(d):
    d.polygon([(16, 100), (40, 18), (112, 34), (100, 110)], fill=BEI, outline=INK)
    line(d, [(16, 100), (40, 18), (112, 34), (100, 110), (16, 100)])
    text(d, "m²", 64, 66, 36, PUR)

@icon("orcamento")
def _(d):
    circ(d, 64, 64, 52, fill=YEL)
    text(d, "R$", 64, 66, 46, INK)

@icon("tempo")
def _(d):
    circ(d, 64, 64, 52, fill=(255, 255, 255, 255))
    line(d, [(64, 64), (64, 26)], w=8); line(d, [(64, 64), (92, 78)], c=RED, w=8)
    circ(d, 64, 64, 6, fill=INK, w=1)

@icon("massa")
def _(d):
    d.polygon([(30, 40), (98, 40), (118, 116), (10, 116)], fill=GRY, outline=INK)
    line(d, [(30, 40), (98, 40), (118, 116), (10, 116), (30, 40)])
    circ(d, 64, 26, 14, fill=None)
    text(d, "kg", 64, 82, 34, (255, 255, 255, 255))

@icon("centro")
def _(d):
    d.polygon([(14, 108), (30, 18), (114, 40), (90, 114)], fill=BEI, outline=INK)
    line(d, [(14, 108), (30, 18), (114, 40), (90, 114), (14, 108)])
    circ(d, 62, 70, 14, fill=None, out=RED, w=6)
    line(d, [(36, 70), (88, 70)], c=RED, w=5); line(d, [(62, 44), (62, 96)], c=RED, w=5)

@icon("cotas")
def _(d):
    rect(d, [30, 10, 118, 78])
    line(d, [(30, 90), (30, 118)], w=4); line(d, [(118, 90), (118, 118)], w=4)
    arrow(d, 64, 104, 116, 104, c=BLU, w=5, h=12); arrow(d, 84, 104, 32, 104, c=BLU, w=5, h=12)
    line(d, [(10, 10), (20, 10)], w=4); line(d, [(10, 78), (20, 78)], w=4)
    arrow(d, 15, 40, 15, 12, c=BLU, w=5, h=10); arrow(d, 15, 48, 15, 76, c=BLU, w=5, h=10)

# ---------------- encaixe ----------------
@icon("enc_ret")
def _(d):
    rect(d, [6, 6, 122, 122], fill=(255, 255, 255, 255), out=BLU, w=5)
    rect(d, [16, 16, 70, 60]); rect(d, [76, 16, 112, 76], fill=BEI2); rect(d, [16, 66, 70, 112], fill=BEI2); rect(d, [76, 82, 112, 112])

@icon("multiplicar")
def _(d):
    for x in (12, 68):
        for y in (12, 68):
            rect(d, [x, y, x + 48, y + 48], fill=BEI if (x + y) % 3 else BEI2, r=8)
    text(d, "x", 64, 64, 40, PUR)

@icon("duplicados")
def _(d):
    rect(d, [14, 30, 80, 96]); rect(d, [34, 14, 100, 80], fill=BEI2)
    line(d, [(82, 82), (118, 118)], c=RED, w=10); line(d, [(118, 82), (82, 118)], c=RED, w=10)

@icon("percurso")
def _(d):
    for (x, y) in [(20, 104), (40, 36), (96, 30), (104, 96)]:
        circ(d, x, y, 13)
    dashed(d, [(20, 104), (40, 36), (96, 30)], c=RED, w=6)
    arrow(d, 96, 30, 104, 82, c=RED, w=6, h=14)

@icon("fechar")
def _(d):
    d.arc([16, 16, 112, 112], 30, 330, fill=INK, width=W)
    line(d, [(105, 40), (105, 88)], c=GRN, w=8)
    node(d, 105, 40, c=GRN); node(d, 105, 88, c=GRN)

@icon("simplificar")
def _(d):
    curve(d, (10, 96), (40, 10), (80, 120), (118, 30), w=7)
    for x, y in [(10, 96), (118, 30)]:
        node(d, x, y)
    text(d, "-", 74, 30, 40, RED)

@icon("agrupar")
def _(d):
    rect(d, [10, 10, 118, 118], r=14)
    circ(d, 42, 44, 14, fill=(255, 255, 255, 255)); circ(d, 86, 44, 14, fill=(255, 255, 255, 255))
    rect(d, [32, 78, 96, 96], fill=(255, 255, 255, 255))
    dashed(d, [(4, 4), (124, 4), (124, 124), (4, 124), (4, 4)], c=PUR, w=4)

@icon("soldar")
def _(d):
    circ(d, 46, 64, 36, fill=BEI, out=None, w=0); circ(d, 82, 64, 36, fill=BEI, out=None, w=0)
    d.arc([10, 28, 82, 100], 40, 320, fill=INK, width=W)
    d.arc([46, 28, 118, 100], 220, 140, fill=INK, width=W)

# ---------------- corte ----------------
@icon("contorno")
def _(d):
    text(d, "A", 64, 66, 76, BLU)
    dashed(d, [(30, 116), (52, 12), (76, 12), (98, 116), (30, 116)], c=MAG, w=6)

@icon("marcas")
def _(d):
    rect(d, [30, 30, 98, 98], fill=(255, 255, 255, 255), out=GRY, w=4)
    for x, y in [(14, 14), (114, 14), (14, 114), (114, 114)]:
        circ(d, x, y, 11, fill=INK, w=1)

@icon("descarte")
def _(d):
    rect(d, [8, 8, 120, 120], fill=(255, 255, 255, 255), out=MAG, w=6)
    for x in (45, 83):
        line(d, [(x, 8), (x, 120)], c=MAG, w=5)
    line(d, [(8, 64), (120, 64)], c=MAG, w=5)
    text(d, "a", 26, 36, 30, INK)

@icon("fatiar")
def _(d):
    for x, y in [(8, 8), (68, 8), (8, 68), (68, 68)]:
        rect(d, [x, y, x + 52, y + 52], fill=BEI)
    text(d, "A1", 34, 34, 22); text(d, "A2", 94, 34, 22)

@icon("gcode")
def _(d):
    rect(d, [8, 14, 120, 98], fill=(40, 45, 60, 255), r=10)
    text(d, "G01", 64, 56, 34, (90, 230, 120, 255))
    rect(d, [44, 100, 84, 118], fill=GRY, w=4)

@icon("dxf")
def _(d):
    rect(d, [18, 8, 110, 120], fill=(255, 255, 255, 255), r=6)
    text(d, "DXF", 64, 52, 32, BLU)
    arrow(d, 40, 94, 88, 94, c=BLU, w=7)

# ---------------- letra caixa ----------------
@icon("letracaixa")
def _(d):
    I["cat_letracaixa"](d)

@icon("led")
def _(d):
    text(d, "A", 64, 68, 110, BEI2)
    for x, y in [(46, 92), (82, 92), (54, 60), (74, 60), (64, 30)]:
        rect(d, [x - 9, y - 6, x + 9, y + 6], fill=YEL, w=3)

@icon("neon")
def _(d):
    curve(d, (12, 96), (30, 0), (70, 140), (116, 30), c=(255, 120, 220, 120), w=22)
    curve(d, (12, 96), (30, 0), (70, 140), (116, 30), c=MAG, w=9)

@icon("furos")
def _(d):
    rect(d, [10, 22, 118, 106], r=16)
    for x in (30, 64, 98):
        circ(d, x, 40, 7, fill=(255, 255, 255, 255), w=4); circ(d, x, 88, 7, fill=(255, 255, 255, 255), w=4)

@icon("ilhoses")
def _(d):
    rect(d, [6, 22, 122, 106], fill=(120, 190, 240, 255))
    for x, y in [(20, 36), (64, 36), (108, 36), (20, 92), (64, 92), (108, 92)]:
        circ(d, x, y, 9, fill=(255, 255, 255, 255), out=GRY, w=5)

# ---------------- geradores ----------------
@icon("malhete")
def _(d):
    pts = [(10, 40)]
    for i, x in enumerate(range(10, 120, 18)):
        y = 40 if i % 2 == 0 else 24
        pts += [(x, y), (x + 18, y)]
    pts += [(118, 40), (118, 114), (10, 114), (10, 40)]
    d.polygon(pts, fill=BEI); line(d, pts, w=6)

@icon("quebracabeca")
def _(d):
    d.rectangle([20, 34, 96, 110], fill=BEI)
    circ(d, 58, 26, 16, fill=BEI, out=None, w=0); circ(d, 104, 72, 16, fill=BEI, out=None, w=0)
    circ(d, 20, 72, 14, fill=(0, 0, 0, 0), out=None, w=0)
    d.arc([42, 10, 74, 42], 180, 360, fill=INK, width=W); d.arc([88, 56, 120, 88], 270, 450, fill=INK, width=W)
    d.arc([6, 58, 34, 86], 270, 450, fill=INK, width=W)
    line(d, [(20, 34), (44, 34)]); line(d, [(72, 34), (96, 34), (96, 58)]); line(d, [(96, 86), (96, 110), (20, 110), (20, 86)]); line(d, [(20, 58), (20, 34)])

@icon("grelha")
def _(d):
    rect(d, [10, 10, 118, 118])
    for k in (46, 82):
        line(d, [(k, 10), (k, 118)], w=5); line(d, [(10, k), (118, k)], w=5)

@icon("circulos")
def _(d):
    for x, y, r in [(36, 38, 26), (88, 30, 18), (92, 84, 28), (40, 96, 18), (64, 64, 9), (112, 112, 8)]:
        circ(d, x, y, r, fill=BEI if r > 15 else BEI2, w=5)

@icon("numerar")
def _(d):
    rect(d, [6, 24, 122, 104], fill=(255, 255, 255, 255), r=10)
    text(d, "123", 64, 64, 40, PUR)

@icon("sequencia")
def _(d):
    rect(d, [20, 6, 108, 54], fill=(255, 255, 255, 255), r=6); text(d, "001", 64, 30, 28, RED)
    rect(d, [20, 70, 108, 118], fill=(255, 255, 255, 255), r=6); text(d, "002", 64, 94, 28, RED)

@icon("caixadelim")
def _(d):
    circ(d, 64, 64, 34)
    dashed(d, [(14, 14), (114, 14), (114, 114), (14, 114), (14, 14)], c=RED, w=6)

# ---------------- edição ----------------
@icon("tamanho")
def _(d):
    rect(d, [34, 34, 94, 94])
    arrow(d, 64, 64, 120, 8, c=PUR, w=7); arrow(d, 64, 64, 8, 120, c=PUR, w=7)

@icon("distribuir")
def _(d):
    for x in (10, 52, 94):
        rect(d, [x, 30, x + 24, 98])
    line(d, [(34, 114), (52, 114)], c=RED, w=5); line(d, [(76, 114), (94, 114)], c=RED, w=5)

@icon("aolongo")
def _(d):
    pts = curve(d, (8, 110), (30, 10), (90, 130), (120, 20), c=GRY, w=5)
    for k in (0, 10, 20, 30, 40):
        x, y = pts[k]
        rect(d, [x - 10, y - 10, x + 10, y + 10], w=4)

@icon("trocar")
def _(d):
    rect(d, [8, 8, 54, 54]); circ(d, 96, 96, 24, fill=BEI2)
    d.arc([30, 30, 110, 110], 280, 350, fill=PUR, width=7)
    arrow(d, 104, 58, 106, 62, c=PUR, w=7, h=14)
    d.arc([18, 18, 98, 98], 100, 170, fill=PUR, width=7)

@icon("espelhar")
def _(d):
    d.polygon([(54, 20), (54, 110), (12, 110)], fill=BEI, outline=INK); line(d, [(54, 20), (54, 110), (12, 110), (54, 20)], w=6)
    d.polygon([(74, 20), (74, 110), (116, 110)], fill=BEI2, outline=INK); line(d, [(74, 20), (74, 110), (116, 110), (74, 20)], w=6)
    dashed(d, [(64, 4), (64, 124)], c=RED, w=5, seg=10)

@icon("angulo")
def _(d):
    line(d, [(14, 112), (118, 112)], w=7); line(d, [(14, 112), (100, 20)], w=7)
    d.arc([-30, 68, 58, 156], 313, 360, fill=RED, width=7)
    text(d, "°", 88, 84, 40, RED)

@icon("arredondar")
def _(d):
    d.rounded_rectangle([14, 14, 140, 140], 50, fill=BEI, outline=INK, width=W)
    circ(d, 64, 64, 5, fill=RED, out=RED, w=1)

@icon("separar")
def _(d):
    text(d, "A", 42, 68, 80, BEI2)
    text(d, "B", 96, 68, 66, PUR)
    dashed(d, [(70, 8), (70, 120)], c=RED, w=4)


@icon("atualizar")
def _(d):
    circ(d, 64, 64, 56, fill=(225, 245, 230, 255), out=GRN, w=5)
    d.arc([28, 28, 100, 100], 200, 340, fill=GRN, width=11)
    d.polygon([(92, 30), (108, 58), (78, 56)], fill=GRN)
    d.arc([28, 28, 100, 100], 20, 160, fill=GRN, width=11)
    d.polygon([(36, 98), (20, 70), (50, 72)], fill=GRN)

@icon("estilos")
def _(d):
    circ(d, 64, 64, 54, fill=(250, 240, 220, 255))
    for (x, y, c) in [(40, 42, MAG), (70, 32, BLU), (92, 56, GRN), (84, 86, YEL)]:
        circ(d, x, y, 11, fill=c, w=3)
    circ(d, 46, 84, 12, fill=(255, 255, 255, 255), w=4)

@icon("config")
def _(d):
    for k in range(8):
        a = k * math.pi / 4
        x, y = 64 + 50 * math.cos(a), 64 + 50 * math.sin(a)
        d.rectangle([x - 11, y - 11, x + 11, y + 11], fill=GRY, outline=INK, width=3)
    circ(d, 64, 64, 42, fill=GRY, w=6)
    circ(d, 64, 64, 16, fill=(255, 255, 255, 255), w=6)

@icon("enc_forma")
def _(d):
    rect(d, [4, 4, 124, 124], fill=(255, 255, 255, 255), out=BLU, w=5)
    circ(d, 44, 80, 32, fill=BEI); circ(d, 44, 80, 16, fill=(255, 255, 255, 255))
    circ(d, 44, 80, 8, fill=BEI2, w=4)
    d.polygon([(70, 20), (116, 20), (116, 66)], fill=BEI2, outline=INK); line(d, [(70, 20), (116, 20), (116, 66), (70, 20)], w=5)
    d.polygon([(82, 116), (120, 116), (120, 78)], fill=BEI, outline=INK); line(d, [(82, 116), (120, 116), (120, 78), (82, 116)], w=5)
    d.arc([60, 30, 100, 70], 120, 300, fill=RED, width=6)

def salvar(name, f):
    im, d = new()
    f(d)
    if name in ("logo", "atualizar") or name.startswith("cat_"):
        im.save(os.path.join(OUT, name + ".ico"), sizes=[(16, 16), (24, 24), (32, 32), (48, 48)], bitmap_format="bmp")
        bg = Image.new("RGBA", (S, S), (240, 240, 240, 255)); bg.alpha_composite(im)
        bg.convert("RGB").resize((32, 32), Image.LANCZOS).save(os.path.join(OUT, name + ".bmp"))
        im.resize((32, 32), Image.LANCZOS).save(os.path.join(OUT, name + ".png"))
    sm = im.resize((24, 24), Image.LANCZOS)
    sm.save(os.path.join(OUT, "m_" + name + ".ico"), sizes=[(24, 24)], bitmap_format="bmp")
    bg = Image.new("RGBA", (24, 24), (255, 255, 255, 255)); bg.alpha_composite(sm)
    bg.convert("RGB").save(os.path.join(OUT, "m_" + name + ".bmp"))
    return im


prev = Image.new("RGBA", (10 * 70, ((len(I) + 9) // 10) * 70), (240, 240, 240, 255))
for k, (n, f) in enumerate(I.items()):
    im = salvar(n, f)
    prev.paste(im.resize((48, 48), Image.LANCZOS), ((k % 10) * 70 + 11, (k // 10) * 70 + 11), im.resize((48, 48), Image.LANCZOS))
prev.save("/home/claude/preview_icones.png")
print(len(I), "icones")

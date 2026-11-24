#!/usr/bin/env python3
"""Erzeugt Lachs-aus-dem-Kaefig.pptx (native, bearbeitbare Formen + Sprechernotizen).
Koordinaten sind die der HTML-Version (1920x1080 px). Aufruf: python3 build_pptx.py"""
import html, pathlib, re
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

root = pathlib.Path(__file__).parent
PX = 6350  # 1 px = 1/144 inch = 6350 EMU  (1920 px = 13,333 in)

C = dict(paper="F3EFE6", ink="1E1D1A", salmon="D98A74", salmon_d="B0583F", sea="23424D",
         teal="4C8A9B", sand="D8CFBB", mute="7D776B", mist="9DBFC8", peach="F1DED3")
# Schriften: Impact/Georgia/Courier New sind auf Windows und macOS vorhanden
FONT = dict(D="Georgia", S="Georgia", M="Courier New")
DS = 0.66  # ruhige Serifenschrift statt Plakat-Versalien -> Anzeigegrößen verkleinern
import os
if os.environ.get("TEST_ANTON"):  # nur für die Layout-Prüfung: Anton mit Impact-äquivalenter Breite
    pass
CUR = {"fg": "ink"}

THEME = {  # Hintergrund, Text, Akzent
    "paper": ("paper", "ink", "salmon_d"),
    "sea": ("sea", "paper", "salmon"),
    "salmon": ("peach", "ink", "salmon_d"),
}

prs = Presentation()
prs.slide_width, prs.slide_height = Emu(1920 * PX), Emu(1080 * PX)
BLANK = prs.slide_layouts[6]


def rgb(k):
    return RGBColor.from_string(C.get(k, k))


def rect(s, x, y, w, h, color):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x * PX)), Emu(int(y * PX)), Emu(max(int(w * PX), 1)), Emu(max(int(h * PX), 1)))
    sh.fill.solid(); sh.fill.fore_color.rgb = rgb(color); sh.line.fill.background(); sh.shadow.inherit = False
    return sh


def text(s, x, y, w, h, paras, size, font="S", color=None, italic=False, align="l", ls=1.0, anchor="t", acc=None, upper=False):
    """paras: str | [paragraph]; paragraph: str | [(text, {color?, italic?, size?})]"""
    color = color or CUR["fg"]
    tb = s.shapes.add_textbox(Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras, str):
        paras = [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT, "c": PP_ALIGN.CENTER}[align]
        p.line_spacing = ls
        runs = [(para, {})] if isinstance(para, str) else para
        for t, o in runs:
            r = p.add_run(); r.text = t.upper() if upper else t
            f = r.font; f.name = FONT[font]
            f.size = Pt(o.get("size", size) * 0.5 * (DS if font == "D" else 1))
            f.italic = o.get("italic", italic); f.bold = False
            f.color.rgb = rgb(o.get("color", color))
    return tb


def D(s, x, y, w, h, paras, size, color=None, **kw):  # Anzeigeschrift, Großbuchstaben
    return text(s, x, y, w, h, paras, size, font="D", color=color, upper=False, ls=kw.pop("ls", 1.0), **kw)


def M(s, x, y, w, txt, color, size=22, align="l"):
    return text(s, x, y, w, 40, txt, size, font="M", color=color, upper=True, align=align)


# Notizen und Labels aus der HTML-Quelle übernehmen
src = (root / "src/deck.html").read_text(encoding="utf-8")
SECTIONS = re.findall(r'<section class="slide (\w+)" data-label="([^"]*)">(.*?)</section>', src, re.S)
NOTES = [html.unescape(re.search(r"<aside>(.*?)</aside>", b, re.S).group(1)) for _, _, b in SECTIONS]
N = len(SECTIONS)


def new(i):
    th, label, _ = SECTIONS[i][0], SECTIONS[i][1], None
    bgk, fg, acc = THEME[th]
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid(); s.background.fill.fore_color.rgb = rgb(bgk)
    CUR["fg"] = fg
    M(s, 120, 52, 1000, label, fg)
    M(s, 1500, 52, 300, f"{i + 1:02d} / {N:02d}", fg, align="r")
    rect(s, 0, 1076, 1920 * (i + 1) / N, 4, "mute")
    s.notes_slide.notes_text_frame.text = NOTES[i]
    return s, fg, acc


def hr(s, x, y, w, color, t=2):
    rect(s, x, y, w, t, color)


def legend(s, x, y, items, fg):
    for k, lab, w in items:
        rect(s, x, y + 2, 20, 20, k); M(s, x + 30, y, w, lab, fg); x += 30 + w + 30


# ---------- 01 Titel
s, fg, acc = new(0)
D(s, 112, 140, 1600, 500, "Lachs", 480, "salmon", ls=0.9)
text(s, 124, 610, 1600, 240, "aus dem Käfig", 190, color="paper", italic=True)
hr(s, 120, 860, 1680, "mute")
text(s, 120, 900, 1500, 80, "Kann Aquakultur die Welt ernähren – am Beispiel Norwegens?", 50, color="paper")

# ---------- 02 Ausgangsfrage
s, fg, acc = new(1)
D(s, 108, 140, 950, 520, "17 %", 430, acc)
text(s, 120, 640, 800, 200, [[("der Menschen decken ihren Eiweißbedarf hauptsächlich über ", {}), ("Fisch", {"italic": True}), (".", {})]], 50)
M(s, 1060, 190, 700, "Fisch weltweit", "mute")
D(s, 1060, 235, 760, 220, ["Wichtigster", "Proteinlieferant"], 92, fg)
text(s, 1060, 460, 740, 70, "noch vor Geflügel und Schwein.", 40)
hr(s, 1060, 570, 740, fg)
text(s, 1060, 600, 740, 120, "Die Weltbevölkerung wächst, die Nachfrage steigt – doch die Meere sind bereits überfischt.", 40, italic=True)
D(s, 1060, 745, 740, 190, "Warum also nicht einfach züchten?", 72, acc)

# ---------- 03 Fang vs Zucht
s, fg, acc = new(2)
D(s, 120, 130, 1600, 270, [[("Der Fang stagniert.", {})], [("Die Zucht wächst.", {"color": acc})]], 118, fg, ls=0.92)
TX, TW = 298, 1502
for yy, yr, tot, f in [(400, "1997", "128,6", 128.6 / 178.5), (575, "2018", "178,5", 1.0)]:
    D(s, 120, yy + 14, 170, 90, yr, 72, fg)
    a = 0.27 if yr == "1997" else 0.46
    wa, wf = TW * f * a, TW * f * (1 - a)
    rect(s, TX, yy, wa, 110, "salmon"); rect(s, TX + wa, yy, wf, 110, "ink")
    D(s, TX + 22, yy + 28, 260, 70, f"{round(a * 100)} %", 58, "ink")
    D(s, TX + wa + 22, yy + 28, 260, 70, f"{100 - round(a * 100)} %", 58, "paper")
    text(s, TX, yy + 116, 600, 40, f"Gesamt {tot} Mio. t", 32, color="mute")
legend(s, TX, 745, [("salmon", "Aquakultur", 190), ("ink", "Fischfang", 170)], fg)
D(s, 120, 880, 330, 130, "95 %", 100, acc)
text(s, 440, 895, 1360, 120, "des Zuwachses seit 1997 stammen aus der Zucht: +47 von +50 Mio. t. Der Fang liegt weiter bei etwa 95 Mio. t.", 42)

# ---------- 04 Wer züchtet
s, fg, acc = new(3)
D(s, 120, 130, 640, 540, [[("China züchtet mehr als ", {}), ("der Rest der Welt zusammen", {"color": acc})]], 100, fg, ls=0.92)
text(s, 120, 720, 560, 240, "58,4 % der weltweit 80,1 Mio. t Zuchtfisch, Krebse und Muscheln (2017). Die zehn größten Länder liefern 88,8 %.", 38, color="mute")
M(s, 800, 128, 1000, "Aquakultur 2017 in Mio. Tonnen", "mute")
rows = [("China", 46.8), ("Indien", 6.2), ("Indonesien", 6.2), ("Vietnam", 3.8), ("Bangladesch", 2.3), ("Ägypten", 1.5),
        ("Norwegen", 1.3), ("Chile", 1.2), ("Myanmar", 1.0), ("Thailand", 0.9)]
for i, (n, v) in enumerate(rows):
    y = 172 + i * 76; hl = n in ("China", "Norwegen")
    text(s, 800, y + 14, 230, 50, n, 32)
    rect(s, 1058, y + 16, max(604 * v / 46.8, 6), 44, "salmon" if hl else "ink")
    D(s, 1690, y + 8, 110, 60, f"{v:.1f}".replace(".", ","), 44, acc if n == "Norwegen" else fg, align="r")

# ---------- 05 Art der Produktion
s, fg, acc = new(4)
D(s, 120, 130, 1680, 140, [[("Gleiche Branche, ", {}), ("zwei Welten", {"color": acc})]], 112, fg)
COL = ["ink", "teal", "salmon", "sand", "mute"]
data = {"China": [54.3, 3.0, 9.7, 31.2, 1.8], "Indien": [87.7, 1.6, 10.6, 0.2, 0], "Vietnam": [66.0, 5.7, 20.7, 7.5, 0.2],
        "Norwegen": [0, 99.8, 0, 0.2, 0], "Chile": [0.3, 70.8, 0, 28.9, 0]}
for i, (n, v) in enumerate(data.items()):
    y = 340 + i * 92
    text(s, 120, y + 18, 240, 50, n, 32)
    x = 378
    for k, pct in enumerate(v):
        w = 1422 * pct / 100
        if w > 0:
            rect(s, x, y + 11, w, 70, COL[k])
            if pct >= 7:
                D(s, x + 16, y + 24, 200, 50, f"{pct:.0f} %", 36, "paper" if COL[k] in ("ink", "teal") else "ink")
        x += w
legend(s, 378, 818, [("ink", "Fisch, Binnenland", 330), ("teal", "Fisch, Küstengewässer", 400), ("salmon", "Krustentiere", 240), ("sand", "Schalentiere", 240)], fg)
text(s, 120, 900, 800, 100, [[("Indien, China: ", {"color": acc}), ("Süßwasserteiche, Karpfen & Co.", {})]], 44)
text(s, 1000, 900, 800, 100, [[("Norwegen: ", {"color": acc}), ("fast nur Fisch im Salzwasser.", {})]], 44)

# ---------- 06 Norwegen im Maßstab
s, fg, acc = new(5)
for i, (big, txt, col) in enumerate([("Platz 7", "unter den Aquakultur-Ländern der Welt: 1,31 Mio. Tonnen (2017)", "ink"),
                                     ("1,6 %", "der Weltproduktion – ein kleiner Anteil …", "ink"),
                                     ("99,8 %", "… der fast komplett im Meer entsteht: Zuchtfisch in Küstengewässern.", "salmon_d")]):
    y = 130 + i * 240
    hr(s, 120, y, 1680, "mute", 1)
    D(s, 120, y + 12, 900, 220, big, 196, col)
    text(s, 1020, y + 50, 780, 170, txt, 44)
hr(s, 120, 850, 1680, "mute", 1)
text(s, 120, 900, 1680, 80, "Norwegen ist keine Großmacht der Menge, sondern der Meeres-Intensivierung – vor allem beim Lachs.", 40, italic=True)

# ---------- 07 Boom
s, fg, acc = new(6)
D(s, 120, 130, 780, 480, [[("Aus dem ", {})], [("Luxusfisch", {"color": acc})], [("wurde Massenware", {})]], 112, fg, ls=0.92)
text(s, 120, 650, 720, 200, "„Die rasante Ausbreitung der Netzkäfige hat den einstigen Luxusfisch demokratisiert.“", 40, italic=True)
text(s, 120, 900, 720, 80, "Preis je Kilo grob: 2001 etwa 22, 2019 etwa 54 NOK.", 32, color="mute")
for x0, lab, vals, nums, mult in [(960, "Verkaufsmenge, Mio. t", (132, 380), ("0,45", "1,3"), "×3"), (1440, "Erlös, Mrd. NOK", (54, 380), ("10", "70"), "×7")]:
    M(s, x0, 140, 400, lab, "mute")
    base = 730
    for k, (h, n, yr) in enumerate(zip(vals, nums, ("2001", "2019"))):
        x = x0 + k * 210
        rect(s, x, base - h, 150, h, "salmon" if k else "ink")
        D(s, x, base - h - 56, 200, 50, n, 44, acc if k else fg)
        D(s, x, base + 10, 200, 50, yr, 44, fg)
    D(s, x0, 810, 300, 140, mult, 120, acc)
text(s, 960, 965, 840, 50, "70 Mrd. NOK sind etwa 6,4 Mrd. € (1 € = 11 NOK)", 32, color="mute")

# ---------- 08 Absatz
s, fg, acc = new(7)
D(s, 120, 130, 660, 520, [[("Fast drei Viertel bleiben ", {}), ("in Europa", {"color": acc})]], 100, fg, ls=0.92)
text(s, 120, 680, 600, 240, "74 % gehen in EU-Staaten und nach Großbritannien, USA und Japan zusammen erreichen nur 8 %.", 38, color="mute")
M(s, 830, 130, 970, "Hauptexportländer 2019, Anteil in %", "mute")
ex = [("Restliche EU", 30, "ink"), ("Sonstige", 18, "mute"), ("Polen", 14, "ink"), ("Frankreich", 9, "ink"), ("Dänemark", 9, "ink"),
      ("Großbritannien", 6, "ink"), ("Spanien", 6, "ink"), ("USA", 5, "salmon"), ("Japan", 3, "salmon")]
for i, (n, v, k) in enumerate(ex):
    y = 175 + i * 80
    text(s, 830, y + 16, 300, 50, n, 32)
    rect(s, 1148, y + 17, 524 * v / 30, 46, k)
    D(s, 1700, y + 10, 100, 60, str(v), 46, acc if k == "salmon" else fg, align="r")

# ---------- 09 Standorte
s, fg, acc = new(8)
D(s, 120, 130, 700, 480, [[("Zucht ", {}), ("dort, wo das Meer ist", {"color": acc})]], 104, fg, ls=0.92)
text(s, 120, 650, 640, 300, "Im Süden und Osten Norwegens, rund um Oslo, gibt es kaum Anlagen. Die Schwerpunkte liegen an der West- und Nordküste.", 40, italic=True)
M(s, 880, 130, 920, "Zuchtanlagen 2019 (Lachs & Regenbogenforelle)", "mute")
for i, (n, v) in enumerate([("Nordland", 188), ("Trøndelag", 178), ("Hordaland", 169), ("Møre og Romsdal", 120), ("Troms", 89), ("Sogn og Fjordane", 81), ("Rogaland", 62)]):
    y = 175 + i * 112
    text(s, 880, y + 18, 340, 50, n, 32)
    rect(s, 1238, y + 6, 562 * v / 188, 68, "salmon" if i < 3 else "mist")
    D(s, 1256, y + 14, 200, 56, str(v), 44, "ink")

# ---------- 10 Lebenslauf
s, fg, acc = new(9)
D(s, 120, 130, 1680, 140, [[("Drei Stationen, ", {}), ("ein Fischleben", {"color": acc})]], 112, fg)
rect(s, 120, 370, 1300, 3, "teal"); M(s, 120, 392, 1200, "Süßwasser · 7 bis 9 Monate", fg)
rect(s, 1440, 370, 360, 3, "ink"); M(s, 1440, 392, 360, "Meerwasser", fg)
steps = [("1 Brutanstalt", "Ei", "40–100 Tage", "Geschlechtsprodukte der Elterntiere werden abgestreift, die Eier reifen (je nach Temperatur)."),
         ("2 Besatzfischfarm", "Larve", "ca. 44 Tage", "Aufzucht der Larven."),
         ("2 Besatzfischfarm", "Parr", "15 Wochen", "Junglachs, an Süßwasser angepasst."),
         ("2 Besatzfischfarm", "Smolt", "4 Wochen", "Umstellung auf das Leben im Salzwasser."),
         ("3 Mastfarm", "Speisefisch", "im Netzgehege", "Bis zu 10 kg Futter pro Tag und Tonne Lachs.")]
for i, (st, big, dur, desc) in enumerate(steps):
    x = 120 + i * 336
    if i:
        rect(s, x - 14, 450, 2, 500, "mute")
    M(s, x, 450, 300, st, "mute", size=20)
    D(s, x, 490, 330, 110, big, 66 if i == 4 else 96, fg)
    D(s, x, 615, 310, 70, dur, 46, acc)
    text(s, x, 700, 300, 240, desc, 26)
text(s, 120, 975, 1680, 60, "Der Zyklus schließt sich: Aus Speisefischen werden wieder Elterntiere.", 40, italic=True)

# ---------- 11 Wild vs Zucht
s, fg, acc = new(10)
D(s, 120, 130, 1680, 140, [[("Was die Farm ", {}), ("der Natur abnimmt", {"color": acc})]], 112, fg)
D(s, 500, 370, 620, 80, "Wildlachs", 64, fg); rect(s, 500, 462, 620, 1, "mist")
D(s, 1175, 370, 625, 80, "Zuchtlachs", 64, acc); rect(s, 1175, 462, 625, 1, "salmon")
tab = [("Süßwasserphase", "Als Parr 1–2 Jahre im Fluss", "Rund 7–9 Monate in Brutanstalt und Besatzfischfarm"),
       ("Wechsel ins Meer", "Der Smolt wandert selbst ab, sobald er die artspezifische Größe hat", "Der Smolt wird auf Salzwasser umgestellt und ins Netzgehege gesetzt"),
       ("Fortpflanzung", "Wanderung flussaufwärts zum Ablaichen", "Geschlechtsprodukte der Elterntiere werden abgestreift"),
       ("Lebensraum im Meer", "Offenes Meer, Wanderfisch", "Netzgehege mit hoher Fischdichte")]
for i, (k, a, b) in enumerate(tab):
    y = 490 + i * 140
    M(s, 120, y + 12, 360, k, "mute", size=20)
    text(s, 500, y, 620, 130, a, 34); text(s, 1175, y, 625, 130, b, 34)

# ---------- 12 Antibiotika
s, fg, acc = new(11)
D(s, 120, 130, 900, 330, [[("Antibiotika: ", {}), ("die Wende", {"color": "salmon_d"})]], 112, fg, ls=0.92)
text(s, 120, 520, 820, 260, [[("Die wichtigste Ursache: ", {}), ("Impfungen", {"italic": True}), (" gegen Infektionskrankheiten – und das trotz massiv wachsender Lachsbestände.", {})]], 44)
D(s, 120, 850, 420, 150, "−98 %", 130, "salmon_d")
text(s, 580, 890, 360, 60, "und mehr", 40, italic=True)
rect(s, 1100, 270, 300, 640, "ink"); rect(s, 1500, 897, 300, 13, "mute")
D(s, 1100, 130, 600, 140, [[("50", {}), (" Tonnen", {"size": 64})]], 150, "ink")
D(s, 1500, 740, 400, 150, [[("<1", {}), (" Tonne", {"size": 64})]], 150, "salmon_d")
D(s, 1100, 930, 300, 70, "1987", 56, "ink"); D(s, 1500, 930, 300, 70, "1990er", 56, "ink")

# ---------- 13 Drei Baustellen
s, fg, acc = new(12)
D(s, 120, 130, 1680, 140, [[("Aber: ", {}), ("drei Baustellen", {"color": acc}), (" bleiben", {})]], 112, fg)
cols = [("01 · Parasiten", "Laus", "8–12 mm groß", "Sie fressen sich in die Haut der Fische, im schlimmsten Fall bis zu Löchern. Die hohe Fischdichte lässt sie sich bestens vermehren. Insektizide helfen immer weniger: Teilweise sind die Läuse resistent."),
        ("02 · Flucht", "Flucht", "Hunderttausende", "Vor allem bei Stürmen reißen Netzgehege, in Extremfällen entkommen Millionen Lachse. Sie mischen sich unter Wildlachse und gefährden deren Genpool."),
        ("03 · Futter", "Futter", "Raubfisch frisst Fisch", "Für die Mast wird noch immer sehr viel Fisch verfüttert (Fischmehl, Fischöl). Neu: mehr Pflanzenöl, Soja, Getreide und Hülsenfrüchte – der Lachs als „Veganer“.")]
for i, (lab, big, sub, body) in enumerate(cols):
    x = 120 + i * 580
    rect(s, x, 380, 520, 1, "mute")
    M(s, x, 405, 500, lab, "mute")
    D(s, x, 445, 520, 150, big, 150, acc)
    D(s, x, 620, 520, 70, sub, 50 if i == 2 else 54, fg)
    text(s, x, 705, 520, 330, body, 32)

# ---------- 14 Pro / Kontra
s, fg, acc = new(13)
D(s, 120, 130, 1680, 140, [[("Blaue Revolution – ", {}), ("Pro und Kontra", {"color": acc})]], 112, fg)
pro = ["Der Wildfang stagniert, jedes Wachstum kommt aus der Zucht", "Fast die Hälfte des Fischs stammt schon aus Aquakultur (46 %)",
       "Antibiotika von 50 t auf unter 1 t: Technik kann Probleme lösen", "Lachs: vom Luxusfisch zum Alltagsprodukt"]
con = ["Läuse vermehren sich in dichten Netzgehegen, Insektizide verlieren Wirkung", "Entkommene Zuchtlachse gefährden den Genpool der Wildlachse",
       "Raubfisch-Mast verbraucht weiterhin viel Fisch (Fischmehl, Fischöl)", "Aus einem Problem im Meer werden Probleme an Land (Soja, Getreide)"]
for x, head, items, col, lk in [(120, "Dafür", pro, fg, "ink"), (1005, "Dagegen", con, acc, "salmon_d")]:
    D(s, x, 340, 795, 90, head, 68, col); rect(s, x, 435, 795, 1, lk)
    for i, t in enumerate(items):
        y = 455 + i * 125
        text(s, x, y + 14, 795, 100, t, 33)
        hr(s, x, y + 118, 795, "mute")

# ---------- 15 Fazit
s, fg, acc = new(14)
D(s, 120, 140, 1680, 560, [[("Es ", {}), ("geht", {"color": acc}), (" – das zeigt die Antibiotika-Wende. Läuse, Fluchten und Futter sind die ", {}), ("offenen Prüfungen.", {"color": acc})]], 126, fg, ls=0.92)
hr(s, 120, 750, 1680, "mute")
text(s, 120, 790, 1000, 140, "Aquakultur ist längst die halbe Antwort auf den Hunger nach Fisch.", 50, color="paper", italic=True)
D(s, 1200, 780, 600, 200, "Nachhaltig – oder nur effizienter?", 80, acc)

out = root / "Lachs-aus-dem-Kaefig.pptx"
prs.save(out)
print("PPTX:", out, out.stat().st_size // 1024, "KB,", len(prs.slides), "Folien")

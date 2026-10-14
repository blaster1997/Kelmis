"""Erzeugt praesentation/svg/08_organigramm.svg (Organigramm EMRIQ, 1169 x 430 px)."""
import os

W, H = 1169, 430
F = 'font-family="Calibri, Carlito, Arial, sans-serif"'
NAVY, BLUE, GREEN, TXT, GREY, LINE = "#002965", "#005AA7", "#95C11F", "#404040", "#7F7F7F", "#CDCDCD"
PAY = "#6E9A12"   # Zahlungsstrom: dunkleres Grün für Linien und Text
out = []


def t(x, y, s, size=15, color=TXT, bold=False, anchor="start"):
    w = ' font-weight="700"' if bold else ""
    out.append(f'<text x="{x}" y="{y}" {F} font-size="{size}" fill="{color}"{w} text-anchor="{anchor}">{s}</text>')


def node(x, y, w, h, title, sub, dark=False):
    fill, stroke = (NAVY, NAVY) if dark else ("#FFFFFF", LINE)
    out.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="8" fill="{fill}" stroke="{stroke}" stroke-width="1.5"/>')
    cx = x + w / 2
    t(cx, y + h / 2 - 4, title, 18, "#FFFFFF" if dark else NAVY, True, "middle")
    t(cx, y + h / 2 + 17, sub, 14, "#FFFFFF" if dark else TXT, False, "middle")


def head(x, y, d):
    """Pfeilspitze an (x, y), Richtung d in u/d/l/r."""
    s = 7
    pts = {"d": [(x, y), (x - s, y - 11), (x + s, y - 11)], "u": [(x, y), (x - s, y + 11), (x + s, y + 11)],
           "l": [(x, y), (x + 11, y - s), (x + 11, y + s)], "r": [(x, y), (x - 11, y - s), (x - 11, y + s)]}[d]
    return " ".join(f"{a},{b}" for a, b in pts)


def flow(points, color, d):
    pts = " ".join(f"{a},{b}" for a, b in points)
    out.append(f'<polyline points="{pts}" fill="none" stroke="{color}" stroke-width="2.5" stroke-linejoin="round"/>')
    x, y = points[-1]
    out.append(f'<polygon points="{head(x, y, d)}" fill="{color}"/>')


# Felder
out.append(f'<rect x="0" y="0" width="812" height="{H}" rx="10" fill="#F2F2F2"/>')
out.append(f'<rect x="840" y="0" width="{W - 840}" height="{H}" rx="10" fill="#E8EFF7"/>')
t(20, 28, "ÖKONOMISCHES ÖKOSYSTEM", 14, GREY, True)
t(860, 28, "FÖRDEREBENE", 14, GREY, True)

# Knoten Ökosystem
node(246, 40, 320, 64, "Entwicklungsgesellschaft", "DG, AGR Kelmis, WFG · SPI für den Gewerbeteil", dark=True)
node(276, 165, 260, 62, "Investoren", "errichten und halten die Gebäude")
node(30, 300, 240, 62, "Unternehmen", "als Mieter")
node(552, 265, 240, 62, "EMRIQ-Haus", "Betreiber der Gründerflächen")
node(552, 362, 240, 56, "Start-ups", "mieten im EMRIQ-Haus")

# Ströme
flow([(381, 104), (381, 163)], BLUE, "d"); t(373, 140, "Infrastruktur, Baurecht", 14, BLUE, True, "end")
flow([(431, 165), (431, 106)], PAY, "u"); t(439, 140, "Erbpachtkanon", 14, PAY, True)
flow([(276, 185), (120, 185), (120, 298)], BLUE, "d"); t(196, 178, "Fläche", 14, BLUE, True, "middle")
flow([(190, 300), (190, 210), (274, 210)], PAY, "r"); t(232, 228, "Miete", 14, PAY, True, "middle")
flow([(536, 185), (712, 185), (712, 263)], BLUE, "d"); t(624, 178, "Fläche (Generalmiete)", 14, BLUE, True, "middle")
flow([(642, 265), (642, 210), (538, 210)], PAY, "l"); t(590, 228, "Miete", 14, PAY, True, "middle")
flow([(642, 327), (642, 360)], BLUE, "d"); t(634, 349, "Fläche", 14, BLUE, True, "end")
flow([(712, 362), (712, 329)], PAY, "u"); t(720, 349, "Miete", 14, PAY, True)

# Legende
out.append(f'<line x1="30" y1="393" x2="62" y2="393" stroke="{BLUE}" stroke-width="2.5"/>')
t(72, 398, "Leistung: Infrastruktur, Baurecht, Fläche", 14, TXT)
out.append(f'<line x1="30" y1="414" x2="62" y2="414" stroke="{PAY}" stroke-width="2.5"/>')
t(72, 419, "Zahlung: Erbpachtkanon, Miete", 14, TXT)

# Förderebene
node(860, 48, 289, 66, "Politik", "DG, Wallonie, EU (EFRE, Interreg)")
node(860, 148, 289, 66, "Unternehmen", "als Förderer und Partner")
node(860, 248, 289, 66, "Forschung", "RWTH, Fraunhofer ILT, ULiège, UM")
t(870, 352, "Fördermittel, Bürgschaften, Programme", 14, NAVY, True)
t(870, 372, "und Ankermieten wirken auf alle", 14, TXT)
t(870, 392, "Stufen des Ökosystems (Hebel 1–12)", 14, TXT)
out.append(f'<polygon points="818,200 834,188 834,212" fill="{BLUE}"/>')

svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">\n'
       + "\n".join(out) + "\n</svg>\n")
path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "svg", "08_organigramm.svg")
open(path, "w", encoding="utf-8").write(svg)
print("geschrieben:", path)

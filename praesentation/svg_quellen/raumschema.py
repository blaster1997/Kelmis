"""Erzeugt praesentation/svg/06_raumschema.svg (schematisches Raumbild Kelmis, 614 x 470 px)."""
import os

W, H = 614, 470
F = 'font-family="Calibri, Carlito, Arial, sans-serif"'
NAVY, BLUE, LBLUE, GREEN, TXT, GREY, LINE, LIGHT = "#002965", "#005AA7", "#009FE3", "#95C11F", "#404040", "#7F7F7F", "#CDCDCD", "#E6E6E6"
out = []


def t(x, y, s, size=15, color=TXT, bold=False, anchor="start"):
    w = ' font-weight="700"' if bold else ""
    out.append(f'<text x="{x}" y="{y}" {F} font-size="{size}" fill="{color}"{w} text-anchor="{anchor}">{s}</text>')


def area(x, y, w, h, fill, opacity=1.0, stroke=None, dash=None, rx=10):
    st = f' stroke="{stroke}" stroke-width="2"' if stroke else ""
    da = f' stroke-dasharray="{dash}"' if dash else ""
    out.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" fill-opacity="{opacity}"{st}{da}/>')


ROAD_Y = 150
# Flächen
area(100, 60, 190, 80, LIGHT)                                    # Ortskern
t(195, 106, "Ortskern Kelmis", 15, TXT, anchor="middle")
area(310, 40, 160, 100, BLUE, 0.15, BLUE, "7 5")                 # Suchraum Nord
t(326, 68, "Suchraum Nord", 16, NAVY, True)
t(326, 90, "Gewerbe, Forschung,")
t(326, 109, "Gründung")
area(300, 190, 155, 88, BLUE, 0.15, BLUE, "7 5")                 # Suchraum Süd
t(316, 217, "Suchraum Süd", 16, NAVY, True)
t(316, 239, "Prüfung gemischte")
t(316, 258, "Nutzung mit Wohnen")
out.append(f'<ellipse cx="140" cy="250" rx="115" ry="48" fill="{GREEN}" fill-opacity="0.2"/>')  # Grünraum
t(140, 247, "Göhltal und Casinoweiher", 15, TXT, anchor="middle")
t(140, 266, "(Natura 2000)", 15, TXT, anchor="middle")

# Lütticher Straße mit Bus 24
out.append(f'<line x1="20" y1="{ROAD_Y}" x2="594" y2="{ROAD_Y}" stroke="{LINE}" stroke-width="10" stroke-linecap="round"/>')
out.append(f'<line x1="300" y1="{ROAD_Y}" x2="590" y2="{ROAD_Y}" stroke="{LBLUE}" stroke-width="2" stroke-dasharray="6 4"/>')
t(20, 132, "Moresnet", 15, GREY)
t(600, 132, "Aachen rd. 8 km", 15, GREY, anchor="end")
t(20, 177, "Lütticher Straße", 15, TXT)
t(594, 177, "Bus 24", 15, LBLUE, anchor="end")
# Museum
out.append(f'<rect x="293" y="{ROAD_Y - 7}" width="14" height="14" fill="{NAVY}" stroke="#FFFFFF" stroke-width="2"/>')
t(300, 177, "Museum Vieille Montagne", 15, NAVY, True, "middle")

# Bahn, Radschnellweg, Bahnhof
out.append(f'<line x1="600" y1="215" x2="505" y2="392" stroke="{BLUE}" stroke-width="2.5"/>')
t(566, 300, "Bahn", 15, BLUE)
out.append(f'<polyline points="580,392 527,350 556,262 600,196" fill="none" stroke="{GREEN}" stroke-width="3" stroke-dasharray="8 5" stroke-linecap="round"/>')
SX, SY = 527, 350
# Verbindungsachse Quartier – Bahnhof
out.append(f'<line x1="455" y1="234" x2="{SX - 6}" y2="{SY - 10}" stroke="{NAVY}" stroke-width="3" stroke-dasharray="8 6"/>')
t(476, 300, "Verbindung rd. 3–4 km", 15, NAVY, True, "end")
t(476, 319, "Shuttle oder Rufbus, Radweg", 15, TXT, anchor="end")
out.append(f'<circle cx="{SX}" cy="{SY}" r="10" fill="{NAVY}" stroke="#FFFFFF" stroke-width="3"/>')
t(SX - 18, 355, "Bahnhof Hergenrath", 15, NAVY, True, "end")
area(372, 365, 128, 27, LIGHT, rx=6)
t(436, 384, "Hochheid (Gewerbe)", 14, TXT, anchor="middle")

# Legende
LY1, LY2 = 418, 440


def leg_rect(x, y, fill, op, stroke=None, dash=None):
    area(x, y - 11, 24, 13, fill, op, stroke, dash, rx=3)


def leg_line(x, y, color, width, dash=None):
    da = f' stroke-dasharray="{dash}"' if dash else ""
    out.append(f'<line x1="{x}" y1="{y - 4}" x2="{x + 24}" y2="{y - 4}" stroke="{color}" stroke-width="{width}"{da}/>')


out.append(f'<line x1="20" y1="398" x2="594" y2="398" stroke="{LIGHT}" stroke-width="1"/>')
leg_rect(20, LY1, BLUE, 0.15, BLUE, "4 3"); t(52, LY1, "Suchraum aus Vorarbeiten", 14)
leg_rect(20, LY2, LIGHT, 1); t(52, LY2, "Ortskern, Gewerbe", 14)
leg_rect(225, LY1, GREEN, 0.25); t(257, LY1, "Grün- und Schutzraum", 14)
leg_line(225, LY2, NAVY, 3, "6 4"); t(257, LY2, "Verbindungsachse", 14)
leg_line(415, LY1, BLUE, 2.5); t(447, LY1, "Bahn mit Halt Hergenrath", 14)
leg_line(415, LY2, GREEN, 3, "6 4"); t(447, LY2, "Radschnellweg (geplant)", 14)
t(20, 464, "Schematische Darstellung ohne Maßstab; Grundlage: Gewerbeflächenstrategie, OBL'EU", 14, GREY)

svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">\n'
       + "\n".join(out) + "\n</svg>\n")
path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "svg", "06_raumschema.svg")
open(path, "w", encoding="utf-8").write(svg)
print("geschrieben:", path)

"""Erzeugt die Präsentation „EMRIQ Kelmis – Leistungskonzept“ auf Basis der Drees-&-Sommer-Vorlage.

Aufruf: python3 praesentation/build_pptx.py
Abbildungen liegen als SVG in praesentation/svg und werden mit PNG-Ersatzbild eingebettet
(PowerPoint zeigt die Vektorgrafik, ältere Programme das PNG).
"""
import copy
import os
import subprocess

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TEMPLATE = os.path.join(HERE, "vorlage", "Vorlage.pptx")
SVG = os.environ.get("EMRIQ_SVG", os.path.join(HERE, "svg"))
PNG = os.environ.get("EMRIQ_PNG", os.path.join(HERE, "png"))
OUT = os.environ.get("EMRIQ_OUT", os.path.join(ROOT, "export", "EMRIQ_Kelmis_Leistungskonzept.pptx"))

FOOTER = "EMRIQ Kelmis – Drees & Sommer – 14.10.2026"
NAVY, BLUE, LBLUE, GREEN = "002965", "005AA7", "009FE3", "95C11F"
TXT, GREY1, GREY2, LINE, WHITE = "404040", "F2F2F2", "E6E6E6", "CDCDCD", "FFFFFF"

os.makedirs(PNG, exist_ok=True)
prs = Presentation(TEMPLATE)
LAYOUT = {l.name.strip(): l for l in prs.slide_layouts}
svg_counter = [0]
svg_parts = {}


# ---------------------------------------------------------------- Hilfsfunktionen
def rgb(h):
    return RGBColor.from_string(h)


def render_png(svg_path, color_swap=None):
    """Rendert ein SVG (optional umgefärbt) als PNG und gibt (svg_bytes, png_path) zurück."""
    data = open(svg_path, encoding="utf-8").read()
    name = os.path.splitext(os.path.basename(svg_path))[0]
    if color_swap:
        for a, b in color_swap.items():
            data = data.replace(a, b)
        name += "_" + "_".join(v.strip("#") for v in color_swap.values())
    tmp_svg = os.path.join(PNG, name + ".svg")
    png = os.path.join(PNG, name + ".png")
    with open(tmp_svg, "w", encoding="utf-8") as f:
        f.write(data)
    if not os.path.exists(png) or os.path.getmtime(png) < os.path.getmtime(tmp_svg):
        subprocess.run(["node", os.path.join(HERE, "render_svg.js"), tmp_svg, png, "3"], check=True, capture_output=True)
    return data.encode("utf-8"), png


def add_svg(slide, svg_path, x, y, w, h, name, color_swap=None):
    if not os.path.exists(svg_path):  # Abbildung noch nicht vorhanden: neutraler Platzhalter
        ph = box(slide, x, y, w, h, GREY1, name=name + " (Platzhalter)")
        text(slide, x + 0.3, y + h / 2 - 0.2, w - 0.6, 0.4, [{"t": "Abbildung folgt", "color": "7F7F7F"}],
             align="c", name=name + " Hinweis")
        return ph
    blob, png = render_png(svg_path, color_swap)
    pic = slide.shapes.add_picture(png, Inches(x), Inches(y), Inches(w), Inches(h))
    pic.name = name
    pic._element.nvPicPr.cNvPr.set("descr", name)
    key = (svg_path, tuple(sorted((color_swap or {}).items())))
    if key not in svg_parts:
        svg_counter[0] += 1
        svg_parts[key] = Part(PackURI(f"/ppt/media/emriq_svg{svg_counter[0]}.svg"), "image/svg+xml",
                              slide.part.package, blob)
    rid = slide.part.relate_to(svg_parts[key], RT.IMAGE)
    blip = pic._element.find(".//" + qn("a:blip"))
    ext_lst = etree.SubElement(blip, qn("a:extLst"))
    ext = etree.SubElement(ext_lst, qn("a:ext"), uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
    svg_blip = etree.SubElement(ext, "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip",
                                nsmap={"asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main"})
    svg_blip.set(qn("r:embed"), rid)
    return pic


def icon(slide, name, x, y, d, circle=BLUE, stroke="#FFFFFF"):
    """Linien-Icon im Farbkreis (Durchmesser d in Zoll)."""
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.name = f"Icon-Kreis {name}"
    fill(c, circle)
    pad = d * 0.2
    add_svg(slide, os.path.join(SVG, "icons", name + ".svg"), x + pad, y + pad, d - 2 * pad, d - 2 * pad,
            f"Icon {name}", color_swap={"#002965": stroke})


def fill(shape, color=None, line=None):
    if color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(color)
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False


def box(slide, x, y, w, h, color=GREY1, name="Fläche", line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.name = name
    fill(s, color, line)
    s.text_frame.text = ""
    return s


def _bullet(p, color):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Emu(Inches(0.2))))
    pPr.set("indent", str(-Emu(Inches(0.2))))
    for tag in ("a:buClr", "a:buFont", "a:buChar", "a:buNone"):
        for e in pPr.findall(qn(tag)):
            pPr.remove(e)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(buClr, qn("a:srgbClr"), val=color)
    etree.SubElement(pPr, qn("a:buFont"), typeface="Arial")
    etree.SubElement(pPr, qn("a:buChar"), char="•")


def text(slide, x, y, w, h, paras, size=14, color=TXT, anchor="t", name="Text", margin=0.0, align="l"):
    """paras: Liste aus Strings oder Dicts {t, size, bold, color, bullet, before, after, runs}."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    m = Inches(margin)
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = m
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, para in enumerate(paras):
        if isinstance(para, str):
            para = {"t": para}
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[para.get("align", align)]
        p.space_before = Pt(para.get("before", 0))
        p.space_after = Pt(para.get("after", 4))
        runs = para.get("runs") or [{"t": para["t"]}]
        for r in runs:
            run = p.add_run()
            run.text = r["t"]
            f = run.font
            f.size = Pt(r.get("size", para.get("size", size)))
            f.bold = r.get("bold", para.get("bold", False))
            f.color.rgb = rgb(r.get("color", para.get("color", color)))
        if para.get("bullet"):
            _bullet(p, para.get("bullet_color", BLUE))
    return tb


def next_id(slide):
    ids = [int(e.get("id")) for e in slide._element.iter() if e.tag == qn("p:cNvPr") and e.get("id", "").isdigit()]
    return max(ids + [1]) + 1


def footer(slide):
    """Fußzeile und Foliennummer wie in der Vorlage (Platzhalter des Layouts)."""
    tree = slide.shapes._spTree
    for kind, idx, name in (("sldNum", 12, "Foliennummernplatzhalter"), ("ftr", 11, "Fußzeilenplatzhalter")):
        if kind == "sldNum":
            body = ('<a:p><a:fld id="{9829449A-3FF3-45CE-9F4F-09BADC0910C6}" type="slidenum"><a:rPr lang="de-DE"/>'
                    '<a:t>‹Nr.›</a:t></a:fld><a:endParaRPr lang="de-DE"/></a:p>')
        else:
            body = f'<a:p><a:r><a:rPr lang="de-DE"/><a:t>{FOOTER.replace("&", "&amp;")}</a:t></a:r></a:p>'
        xml = (f'<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
               f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
               f'<p:nvSpPr><p:cNvPr id="{next_id(slide)}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
               f'<p:nvPr><p:ph type="{kind}" sz="quarter" idx="{idx}"/></p:nvPr></p:nvSpPr><p:spPr/>'
               f'<p:txBody><a:bodyPr/><a:lstStyle/>{body}</p:txBody></p:sp>')
        tree.append(etree.fromstring(xml))


def new_slide(title, subtitle, notes):
    s = prs.slides.add_slide(LAYOUT["Only Title"])
    for ph in s.placeholders:
        if ph.placeholder_format.type is not None and ph.placeholder_format.idx == 0:
            ph.text_frame.text = title
        elif ph.placeholder_format.idx == 14:
            ph.text_frame.text = subtitle
    footer(s)
    s.notes_slide.notes_text_frame.text = notes
    return s


def ap_tag(slide, x, y, label, color=NAVY, w=0.62):
    t = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.27))
    t.name = f"Etikett {label}"
    fill(t, color)
    tf = t.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = rgb(WHITE)
    return t


# ---------------------------------------------------------------- Vorlage vorbereiten
# Musterfolien 2 und 3 entfernen, Titelfolie behalten
sld_ids = prs.slides._sldIdLst
for sld_id in list(sld_ids)[1:]:
    prs.part.drop_rel(sld_id.rId)
    sld_ids.remove(sld_id)

# ---------------------------------------------------------------- Folie 1: Titel
s1 = prs.slides[0]
for sp in list(s1.placeholders):
    pf = sp.placeholder_format
    if pf.idx == 15:  # leerer Bildplatzhalter
        sp._element.getparent().remove(sp._element)
    elif pf.idx == 0:
        sp.text_frame.paragraphs[0].runs[0].text = "EMRIQ Kelmis"
    elif pf.idx == 1:
        ps = sp.text_frame.paragraphs
        ps[0].runs[0].text = "Städtebauliches Entwicklungskonzept mit Umsetzungsplan – Leistungskonzept"
        ps[1].runs[0].text = "Ministerium der Deutschsprachigen Gemeinschaft  ·  Fachbereich Raumordnung, Wohnungswesen und Energie"
        ps[2].runs[0].text = "Gesprächstermin  ·  14. Oktober 2026"
    elif pf.idx == 11:
        sp.text_frame.paragraphs[0].runs[0].text = FOOTER
s1.notes_slide.notes_text_frame.text = (
    "Begrüßung und Vorstellung. Ziel des Termins: unser Verständnis der Aufgabe und das Leistungskonzept vorstellen, "
    "Fragen des Ministeriums aufnehmen und offene Punkte für das Angebot klären. Dauer rund 30 Minuten Vortrag, "
    "danach Diskussion.")

# ---------------------------------------------------------------- Folie 2: Ausgangslage
s = new_slide("Ausgangslage und Aufgabe",
              "Ein Innovationsquartier mit klarem Profil – verankert in Kelmis und der Euregio",
              "Kelmis liegt zwischen den Hochschulstandorten Aachen, Lüttich und Maastricht, rund 8 km von Aachen, "
              "mit Bahnhalt Hergenrath und Buslinie 24. Das Kandidatengebiet für das Einstein-Teleskop grenzt an. "
              "Die DG will hier ein thematisch profiliertes Quartier, kein Gewerbegebiet. Die sechs Erfolgskriterien "
              "rechts sind unser Maßstab für die gesamte Bearbeitung.")
add_svg(s, os.path.join(SVG, "01_euregio_lage.svg"), 0.57, 1.75, 5.896, 4.896, "Lagekarte Euregio")
text(s, 6.95, 1.80, 5.8, 1.55, [
    {"t": "Was die Deutschsprachige Gemeinschaft beabsichtigt", "size": 16, "bold": True, "color": NAVY, "after": 6},
    {"t": "Ein Quartier, das Forschung, innovative Unternehmen, Produktion, Gründungen, Wohnen und Lebensqualität "
          "verbindet und die Brückenfunktion der DG zwischen Belgien, Nordrhein-Westfalen und den Hochschulen in "
          "Aachen, Lüttich und Maastricht stärkt."},
], name="Text Absicht")
text(s, 6.95, 3.50, 5.8, 0.35, [{"t": "Woran wir den Erfolg messen", "size": 16, "bold": True, "color": NAVY}],
     name="Überschrift Erfolgskriterien")
crit = ["Schlüssige thematische Profilierung mit Alleinstellung in der Euregio",
        "Nachgewiesene Bedarfe und klar beschriebene Zielgruppen",
        "Robustes Struktur- und Nutzungskonzept mit Entwicklungsspielräumen",
        "Wirtschaftlich und organisatorisch tragfähiges Umsetzungsmodell",
        "Priorisierte Roadmap mit Kosten, Folgekosten und Zuständigkeiten",
        "Standortstory in Deutsch, Französisch und Englisch"]
for i, c in enumerate(crit):
    y = 3.95 + i * 0.44
    add_svg(s, os.path.join(SVG, "icons", "haken.svg"), 6.95, y + 0.02, 0.26, 0.26, f"Haken {i + 1}",
            color_swap={"#002965": "#" + GREEN})
    text(s, 7.35, y, 5.4, 0.36, [c], name=f"Kriterium {i + 1}", anchor="m")

# ---------------------------------------------------------------- Folie 3: Ansatz
s = new_slide("Unser Ansatz",
              "Strategie, Planung und Umsetzung in einem Prozess – aufbauend auf den vorhandenen Grundlagen",
              "Vier Grundsätze prägen unser Vorgehen. Wir nutzen die vorhandenen Studien als Startpunkt und investieren "
              "das Budget in Fokussierung, Zahlen und Umsetzung. Die thematische Entscheidung bleibt bei der DG; wir "
              "liefern eine transparente Entscheidungsgrundlage. Besonders ist die wirtschaftliche Betrachtung aus "
              "Sicht von Eigentümer, Käufern und Investoren.")
cards = [
    ("grundlagen", "Grundlagen nutzen",
     "Fachgutachten Raumordnung, Gewerbeflächenstrategie, OBL'EU und MORO werden zu Beginn zusammengeführt. "
     "Wir ergänzen gezielt, statt neu zu analysieren."),
    ("fokus", "Fokussieren",
     "Ein Thema statt Campusstrategie. Schwerpunkte werden transparent bewertet; die DG entscheidet auf "
     "nachvollziehbarer Grundlage."),
    ("rechnen", "Rechnen",
     "Das Quartier wird aus Sicht von Eigentümer, Käufern und Investoren durchgerechnet: Exit-Szenarien, "
     "Grundstückswerte, Finanzierungsbedarf."),
    ("umsetzen", "Umsetzen und vermitteln",
     "Roadmap mit Zuständigkeiten und Meilensteinen sowie eine dreisprachige Standortstory für Investoren "
     "und Fachkräfte."),
]
cw, gap = 2.82, 0.30
for i, (ic, head, body) in enumerate(cards):
    x = 0.57 + i * (cw + gap)
    box(s, x, 1.85, cw, 3.45, GREY1, name=f"Karte {head}")
    icon(s, ic, x + 0.25, 2.10, 0.75)
    text(s, x + 0.25, 2.95, cw - 0.5, 0.55, [{"t": head, "size": 16, "bold": True, "color": NAVY}], anchor="b",
         name=f"Titel {head}")
    text(s, x + 0.25, 3.62, cw - 0.5, 1.6, [body], name=f"Text {head}")
band = box(s, 0.57, 5.55, 12.18, 1.05, NAVY, name="Team")
icon(s, "team", 0.82, 5.70, 0.75, circle=BLUE)
text(s, 1.85, 5.62, 10.7, 0.92, [
    {"t": "Interdisziplinäres Kernteam: Stadt- und Quartiersentwicklung · Immobilienökonomie · Projektmanagement · "
          "Nachhaltigkeit · Mobilität · Kommunikation", "bold": True, "color": WHITE, "after": 4},
    {"t": "Projektleitung und Team mit Deutsch, Französisch und Englisch · persönlicher Ansprechpartner, "
          "bei Bedarf vor Ort in Eupen", "color": WHITE},
], anchor="m", name="Text Team")

# ---------------------------------------------------------------- Folie 4: Leistungsumfang
s = new_slide("Leistungsumfang im Überblick",
              "Acht Arbeitspakete in drei Phasen, begleitet von einem durchgängigen Abstimmungsprozess",
              "Die Arbeitspakete folgen der Logik des Lastenhefts: erst Grundlagen und Analyse, dann Profil und "
              "Konzept, zuletzt Wirtschaftlichkeit und Umsetzung. AP 8 begleitet alle Phasen. Jedes Paket endet mit "
              "einem prüfbaren Ergebnis und einem Meilenstein mit Freigabe durch die DG.")
add_svg(s, os.path.join(SVG, "04_leistungsarchitektur.svg"), 0.57, 1.75, 12.177, 4.854, "Leistungsarchitektur")

# ---------------------------------------------------------------- Folie 5: AP 1–3
s = new_slide("AP 1–3: Grundlagen, Analyse und Themenprofil",
              "Vom gemeinsamen Zielsystem zur begründeten thematischen Ausrichtung",
              "AP 1 schafft eine gemeinsame Arbeitsgrundlage und einen Kriterienrahmen. AP 2 analysiert Standort, "
              "Mobilität und Markt einschließlich möglicher Effekte des Einstein-Teleskops. In AP 3 wird das Thema "
              "in vier Stufen eingegrenzt; die Entscheidung fällt im ersten Workshop. Rückfrage an das Ministerium: "
              "Gibt es bereits eine politische Themenpräferenz?")
aps = [
    ("AP 1", "Projektstart, Grundlagen und Zielsystem",
     ["Auswertung vorhandener Gutachten, Strategien und Geodaten",
      "Zielsystem, Betrachtungsraum und Kriterienrahmen"],
     "Abgestimmte Arbeitsgrundlage und Bearbeitungsplan"),
    ("AP 2", "Standort-, Markt- und Potenzialanalyse",
     ["Standort, Mobilität und Wohnungsmarkt",
      "Vergleichsstandorte, Nachfrage, Effekte des Einstein-Teleskops"],
     "Standortsteckbrief mit SWOT und Entwicklungshebeln"),
    ("AP 3", "Themenprofil, Zielgruppen und Ecosystem",
     ["Bewertung der Schwerpunkte, Empfehlung mit Alternativen",
      "Zielgruppen, Anforderungen, Partner- und Betreibermodell"],
     "Begründetes Themen- und Zielgruppenprofil"),
]
for i, (tag, head, bullets, result) in enumerate(aps):
    y = 1.80 + i * 1.62
    ap_tag(s, 0.57, y + 0.02, tag)
    text(s, 1.35, y, 5.4, 0.32, [{"t": head, "size": 15, "bold": True, "color": NAVY}], name=f"Titel {tag}")
    text(s, 1.35, y + 0.36, 5.4, 0.78, [{"t": b, "bullet": True, "after": 2} for b in bullets], name=f"Inhalt {tag}")
    text(s, 1.35, y + 1.12, 5.4, 0.3, [{"runs": [{"t": "Ergebnis: ", "bold": True, "color": GREEN},
                                                 {"t": result}]}], name=f"Ergebnis {tag}")
add_svg(s, os.path.join(SVG, "05_themenprofil_trichter.svg"), 7.05, 1.75, 5.698, 4.854, "Trichter Themenprofil")

# ---------------------------------------------------------------- Folie 6: AP 4–5
s = new_slide("AP 4–5: Entwicklungskonzept, Nachhaltigkeit und Mobilität",
              "Varianten, Vorzugsvariante und Phasierung – verknüpft mit Ortskern und Bahnhof Hergenrath",
              "Das Raumbild zeigt die Suchräume aus den Vorarbeiten; der Betrachtungsraum wird in AP 1 gemeinsam "
              "festgelegt. Entscheidend ist die Verbindung zum Bahnhof Hergenrath, weil die Erreichbarkeit mit Bus und "
              "Bahn heute der größte Engpass ist. Die Vorzugsvariante wird mit der DG verbindlich festgelegt und ist "
              "Grundlage für Wirtschaftlichkeit und Umsetzungsplan.")
add_svg(s, os.path.join(SVG, "06_raumschema.svg"), 0.57, 1.75, 6.396, 4.896, "Raumschema Kelmis")
panels = [
    ("AP 4", "Städtebauliches Entwicklungs- und Nutzungskonzept",
     ["Räumlich-funktionales Leitbild und Flächenprogramm", "Städtebau, Freiraum, Erschließung, Infrastruktur",
      "Variantenvergleich und Vorzugsvariante", "Phasierung, Flexibilität, erste Aktivierungsmaßnahmen"]),
    ("AP 5", "Nachhaltigkeit, Mobilität und Resilienz",
     ["Ziele für Energie, Klima, Ressourcen, Biodiversität und soziale Qualität",
      "Multimodale Erreichbarkeit mit weniger Autoverkehr", "Klimaangepasste Freiräume, Regenwasser, zirkuläres Bauen",
      "Messbare Indikatoren für Planung und Umsetzung"]),
]
for i, (tag, head, bullets) in enumerate(panels):
    y = 1.80 + i * 2.45
    box(s, 7.25, y, 5.5, 2.30, GREY1, name=f"Fläche {tag}")
    ap_tag(s, 7.45, y + 0.2, tag)
    text(s, 8.20, y + 0.16, 4.4, 0.62, [{"t": head, "size": 15, "bold": True, "color": NAVY}], name=f"Titel {tag}")
    text(s, 7.45, y + 0.80, 5.15, 1.45, [{"t": b, "bullet": True, "after": 3} for b in bullets], name=f"Inhalt {tag}")

# ---------------------------------------------------------------- Folie 7: AP 6
s = new_slide("AP 6: Wirtschaftlichkeit, Business Case und Finanzierung",
              "Das Quartier aus Sicht von Eigentümer, Käufern und Investoren rechnen",
              "Basisstrategie ist die Entwicklung bis zum Planungsrecht mit anschließender Erschließung und parallelem "
              "Verkauf der Grundstücke. So können Erlöse früh fließen und einen Teil der Vorleistungen refinanzieren. "
              "Die Grundstückspreise schätzen wir über die Residualwertmethode aus Sicht der späteren Projektentwickler. "
              "Wir vergleichen verschiedene Exit-Zeitpunkte und empfehlen Grundstücksgrößen, die für Käufer "
              "finanzierbar sind. Rechtliche Bewertungen und Verkehrswertgutachten sind nicht Teil der Leistung.")
text(s, 0.57, 1.78, 6.0, 0.3, [{"t": "Basisstrategie und Exit-Zeitpunkte", "size": 15, "bold": True, "color": NAVY}],
     name="Überschrift Exit")
add_svg(s, os.path.join(SVG, "07_wertentwicklung_exit.svg"), 0.57, 2.12, 6.0, 3.198, "Wertentwicklung Exit")
text(s, 6.85, 1.78, 5.9, 0.3, [{"t": "Grundstückswert nach Residualwertmethode", "size": 15, "bold": True,
                                "color": NAVY}], name="Überschrift Residualwert")
add_svg(s, os.path.join(SVG, "07_residualwert.svg"), 6.85, 2.12, 5.896, 3.198, "Residualwert")
cards = [("Exit-Szenarien", "Verkauf nach Planungsrecht, parallel zur Erschließung oder nach Vermarktungsstart"),
         ("Slot-Sizes", "Grundstücksgrößen, die finanzierbar und kapitalmarktfähig sind"),
         ("Parzellierung", "Eigenständig erschließbare, entwickelbare und vermarktbare Grundstücke"),
         ("Öffentliche Hand", "Kosten, Erlöse und Finanzierungsbedarf, Folgekosten und Einnahmen")]
for i, (head, body) in enumerate(cards):
    x = 0.57 + i * (cw + gap)
    box(s, x, 5.42, cw, 1.22, GREY1, name=f"Karte {head}")
    text(s, x + 0.18, 5.50, cw - 0.36, 1.08, [{"t": head, "size": 14, "bold": True, "color": NAVY, "after": 2},
                                              {"t": body, "size": 13}], name=f"Text {head}")

# ---------------------------------------------------------------- Folien 8–12: Struktur und Förderung
# Grundlage: Arbeitsmaterial „Organigramm Innovationscampus“ (Organigramm, Stakeholder, Förderhebel, Vorbilder),
# übertragen auf Kelmis und den belgischen bzw. wallonischen Förderrahmen.
FUNDER = {"Politik": NAVY, "Unternehmen": BLUE, "Forschung": GREEN}


def chips(slide, x, y, names):
    for n in names:
        w = 0.12 + 0.075 * len(n)
        c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.22))
        c.name = f"Träger {n}"
        fill(c, FUNDER[n])
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = n
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = rgb(WHITE)
        x += w + 0.08


def num(slide, x, y, n, d=0.38, color=BLUE, size=13):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.name = f"Nummer {n}"
    fill(c, color)
    tf = c.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(n)
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = rgb(WHITE)


# Folie 8: Organigramm
s = new_slide("Strukturierung: Organigramm des Quartiers",
              "Boden und Erschließung in öffentlicher Hand – Bau durch Investoren – Betrieb durch einen Hub-Betreiber",
              "Vorschlag zur Diskussion, kein Ergebnis. Die Entwicklungsgesellschaft – zunächst als Projektallianz aus DG, "
              "AGR Kelmis und WFG, für den Gewerbeteil mit der SPI – schafft Baurecht und Erschließung und vergibt die "
              "Baufelder in Erbpacht. Investoren bauen und halten die Gebäude und vermieten an Unternehmen sowie per "
              "Generalmietvertrag an den Betreiber des EMRIQ-Hauses, der die Flächen kurzfristig an Start-ups weitergibt. "
              "Politik, Unternehmen und Forschung wirken über Fördermittel, Bürgschaften, Programme und Ankermieten.")
add_svg(s, os.path.join(SVG, "08_organigramm.svg"), 0.57, 1.80, 12.177, 4.479, "Organigramm EMRIQ")
text(s, 0.57, 6.36, 12.18, 0.3, [{"runs": [
    {"t": "Grundsätze: ", "bold": True, "color": NAVY},
    {"t": "Steuerung über Erbpacht statt Verkauf · Immobilie und Betrieb getrennt · Gründerflächen befristet, "
          "Anschlussflächen im Quartier"}], "size": 13}], name="Grundsätze Organigramm")

# Folie 9: Interessen der Stakeholder
s = new_slide("Interessen der Stakeholder",
              "Was jede Rolle erreichen will, wo Zielkonflikte liegen und welche Förderhebel ansetzen",
              "Die Tabelle zeigt, warum sich die Akteure beteiligen und wo es haken kann. Die Nummern verweisen auf die "
              "Förderhebel der beiden folgenden Folien. Zentraler Zielkonflikt der Entwicklungsgesellschaft: Ein hoher "
              "Erbpachtkanon schreckt Investoren ab, ein niedriger gefährdet die Refinanzierung. Beim Betreiber des "
              "EMRIQ-Hauses ist es die Fristeninkongruenz zwischen langer Generalmiete und kurzen Start-up-Verträgen.")
rows = [
    ("Entwicklungsgesellschaft", "DG, AGR Kelmis, WFG, SPI",
     "Quartier zügig und mit klarem Profil entwickeln; Nutzungen über Erbpacht steuern; Erschließung refinanzieren",
     "Hohe Vorleistung vor den ersten Einnahmen; hoher Kanon schreckt Investoren ab, niedriger gefährdet die Refinanzierung",
     "1 2 3 4"),
    ("Investoren", "errichten und halten die Gebäude",
     "Rendite und Werterhalt; Ankermieter und lange Laufzeiten; kalkulierbarer Kanon; Drittverwendbarkeit",
     "Leerstand bei Speziallaboren; Start-ups mit geringer Bonität; kurze Restlaufzeit der Erbpacht",
     "1 5 6 8"),
    ("Unternehmen", "als Mieter",
     "Nähe zu Fachkräften, Forschung und Start-ups; flexible Büro-, Labor- und Testflächen",
     "Mieten über dem Niveau des Umlands; Bindung ohne messbaren Innovationsnutzen",
     "7 10"),
    ("EMRIQ-Haus", "Betreiber der Gründerflächen",
     "Hohe Auslastung und Kostendeckung; gute Auswahl der Start-ups; Partner gewinnen",
     "Lange Generalmiete, kurze Start-up-Verträge; Leerstand und Mietausfälle in der Anlaufphase",
     "8 9 10 12"),
    ("Start-ups", "mieten im EMRIQ-Haus",
     "Günstige, flexible Flächen; Zugang zu Kapital, Pilotkunden und Laboren",
     "Knappe Liquidität; Anschlussflächen beim Wachstum; Abhängigkeit von Förderzyklen",
     "11 12"),
    ("Politik", "Gemeinde, DG, Wallonie, EU",
     "Arbeitsplätze, Einnahmen, Sichtbarkeit; hohe Wirkung je Euro; beihilferechtssichere Förderung",
     "Mitnahmeeffekte; lange Haushaltsbindung (Gemeinde unter Aufsicht); politisches Risiko bei schwachem Anlauf",
     "1 2 3 4 8 9 11"),
    ("Unternehmen", "als Förderer und Partner",
     "Früher Zugang zu Technologien und Start-ups; Kooperation mit der Forschung; Fachkräfte",
     "Unklarer Nutzen; Fragen des geistigen Eigentums; Abgrenzung zu eigenen Programmen",
     "5 10 11"),
    ("Forschung", "RWTH, Fraunhofer ILT, ULiège, UM",
     "Transfer und Ausgründungen; Industrieprojekte und Drittmittel; Reallabore",
     "Begrenzte Kapazität für Transfer; Spannung zwischen Grundlagen- und Auftragsforschung",
     "5 6 7"),
]
cols = [2.35, 4.15, 4.35, 1.33]
tbl = s.shapes.add_table(len(rows) + 1, 4, Inches(0.57), Inches(1.80), Inches(sum(cols)), Inches(4.75)).table
for j, w in enumerate(cols):
    tbl.columns[j].width = Inches(w)
tbl.rows[0].height = Inches(0.34)
for i in range(1, len(rows) + 1):
    tbl.rows[i].height = Inches(0.53)


def cell(c, paras, bg, color=TXT, size=11):
    c.fill.solid()
    c.fill.fore_color.rgb = rgb(bg)
    c.margin_left = c.margin_right = Inches(0.08)
    c.margin_top = c.margin_bottom = Inches(0.04)
    c.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = c.text_frame
    tf.word_wrap = True
    for k, pr in enumerate(paras):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(0)
        r = p.add_run()
        r.text = pr["t"]
        r.font.size = Pt(pr.get("size", size))
        r.font.bold = pr.get("bold", False)
        r.font.color.rgb = rgb(pr.get("color", color))


for j, h in enumerate(["Stakeholder", "Kerninteressen", "Risiken und Zielkonflikte", "Förderhebel"]):
    cell(tbl.cell(0, j), [{"t": h, "bold": True}], NAVY, WHITE, 12)
for i, (name, sub, interest, risk, levers) in enumerate(rows, start=1):
    bg = GREY1 if i % 2 else WHITE
    if i == 6:  # Förderebene beginnt
        bg = "E8EFF7"
    elif i > 6:
        bg = "E8EFF7" if i % 2 == 0 else "F3F7FB"
    cell(tbl.cell(i, 0), [{"t": name, "bold": True, "color": NAVY, "size": 12}, {"t": sub, "size": 10, "color": "595959"}], bg)
    cell(tbl.cell(i, 1), [{"t": interest}], bg)
    cell(tbl.cell(i, 2), [{"t": risk}], bg)
    cell(tbl.cell(i, 3), [{"t": levers, "bold": True, "color": BLUE, "size": 12}], bg)
text(s, 0.57, 6.64, 12.18, 0.22, [{"runs": [
    {"t": "Grau: ökonomisches Ökosystem · Blau hinterlegt: Förderebene · ", "color": "7F7F7F"},
    {"t": "Nummern = Förderhebel auf den folgenden Folien (Ökosystem: überzeugt den Akteur; Förderebene: finanziert oder trägt er)",
     "color": "7F7F7F"}], "size": 11}], name="Lesehilfe Stakeholder")

# Folien 10 und 11: Förderhebel
levers_a = [
    (1, "GRUNDSTÜCK UND INFRASTRUKTUR", "Gestaffelter Erbpachtkanon",
     "Niedrig in der Anlaufphase, danach steigend und an eine Nutzungsbindung gekoppelt. Senkt die Kosten der Investoren.",
     "Gestaffelte Erbpacht nach Buch 3 ZGB; im SPI-Gebiet ist der Verkehrswert die Untergrenze",
     ["Politik"], "Vorbild: Zürich – Baurecht und Unterbaurecht"),
    (2, "GRUNDSTÜCK UND INFRASTRUKTUR", "Infrastruktur nicht umlegen",
     "Die öffentliche Hand trägt Erschließung, Plätze, Mobilität und Entwicklungsmanagement, ohne sie auf Kanon oder Preis umzulegen.",
     "SPI-Erschließung nach Dekret 2017 (65 % plus bis 20 Punkte), Infrastrukturplan der DG, EFRE",
     ["Politik"], "Vorbild: Zürich 85,7 Mio. CHF; Cornell Tech 100 Mio. $"),
    (3, "GRUNDSTÜCK UND INFRASTRUKTUR", "Grunderwerb und Brachen fördern",
     "Zuschüsse für den Ankauf und die Sanierung untergenutzter Flächen senken die Kosten der Baulandentwicklung.",
     "Dekret 2017: Grunderwerb in Wohn- und Sanierungsgebieten 50 %; Brachen gezielt sanieren",
     ["Politik"], "Im Einzelfall rechtlich prüfen"),
    (4, "GRUNDSTÜCK UND INFRASTRUKTUR", "Refinanzierung aus Wertsteigerung",
     "Infrastruktur wird vorfinanziert und aus späteren Rückflüssen zurückgezahlt – Vorleistung statt verlorener Zuschuss.",
     "Revolvierender Bodenfonds mit Rückflüssen aus Erbpacht und Verkäufen; städtebauliche Lasten",
     ["Politik"], "Vorbild: Cortex St. Louis, Zürich"),
    (5, "GEBÄUDE", "Ankermieter und Vorvermietung",
     "Ein bonitätsstarker Mieter macht das Gebäude für Investoren finanzierbar.",
     "Generalmietvertrag von WFG oder DG für das EMRIQ-Haus; Museum, Forschungsaußenstelle oder Projektbüro Einstein-Teleskop als Ankermieter",
     ["Politik", "Forschung", "Unternehmen"], "Vorbild: Block71 Singapur, High Tech Campus Eindhoven"),
    (6, "GEBÄUDE", "Gemeinsame Forschungsinfrastruktur",
     "Zuschüsse oder Bürgschaften für Labore, Messtechnik und Geräte, die Investoren ungern finanzieren – für alle Mieter nutzbar.",
     "Schwingungsarme Labore und Messplätze über EFRE und Interreg, mit RWTH oder Fraunhofer ILT",
     ["Politik", "Forschung"], "Vorbild: Switzerland Innovation, Eindhoven"),
]
levers_b = [
    (7, "NUTZER", "F&E-Kooperation und Steueranreize",
     "Verbundprojekte der Mieter mit Forschung und Start-ups; steuerliche Förderung eigener Forschung.",
     "Interreg-Verbundprojekte; belgischer Innovationsabzug und Teilbefreiung vom Berufssteuervorabzug für Forschende",
     ["Politik", "Forschung", "Unternehmen"], "Vorbild: RWTH Aachen Campus"),
    (8, "BETRIEB", "Mietgarantie in der Anlaufphase",
     "Die öffentliche Hand sichert den Generalmietvertrag des Betreibers für die ersten Jahre ab und schließt so die Fristenlücke.",
     "Befristete, gedeckelte Garantie von DG oder WFG gegenüber dem Investor des EMRIQ-Hauses",
     ["Politik"], "Vorbild: Bürgschaftsmodell Switzerland Innovation"),
    (9, "BETRIEB", "Investitions- und Betriebszuschüsse",
     "Zuschüsse für Ausbau, Ausstattung und Anlaufkosten des Gründerzentrums.",
     "Dekret 2017: Gründerzentrum bis 80 % (gedeckelt); EFRE 2028–2034; Betriebszuschuss der DG in der Anlaufphase",
     ["Politik"], "Programmierung EFRE 2027 nutzen"),
    (10, "BETRIEB", "Partner- und Mitgliedschaftsmodelle",
     "Unternehmen zahlen Mitgliedschaften oder betreiben eigene Programme im Hub und finanzieren Betrieb und Angebote mit.",
     "Partner wie HAHN Automation und Aachener Unternehmen mit eigenen Programmen",
     ["Unternehmen"], "Vorbild: RWTH Campus, Station F Paris"),
    (11, "GRÜNDUNG", "Direktförderung der Start-ups",
     "Zuschüsse, Inkubatorprogramme und Wagniskapital direkt an die Gründungen.",
     "Wagniskapital über Noshaq und Wallonie Entreprendre; EIC Accelerator; Programme der WFG",
     ["Politik", "Unternehmen"], "Vorbild: Inkubatoren Israel, Cornell Tech"),
    (12, "GRÜNDUNG", "Einstiegsmieten mit Verweildauer",
     "Gestaffelte oder bezuschusste Mieten, befristet, damit Flächen rotieren und Wachstum im Quartier bleibt.",
     "Verträge über drei bis fünf Jahre im EMRIQ-Haus, danach Anschlussflächen auf den Gewerbebaufeldern",
     ["Politik", "Unternehmen"], "Vorbild: Block71 Singapur (3 Jahre, Firmen unter 5 Jahren)"),
]


def lever_slide(title, subtitle, notes, levers):
    s = new_slide(title, subtitle, notes)
    cw3, ch = 3.86, 2.30
    for k, (n, stage, head, body, kelmis, who, ref) in enumerate(levers):
        x = 0.57 + (k % 3) * (cw3 + 0.30)
        y = 1.80 + (k // 3) * (ch + 0.18)
        box(s, x, y, cw3, ch, GREY1, name=f"Hebel {n}")
        num(s, x + 0.15, y + 0.15, n)
        text(s, x + 0.65, y + 0.12, cw3 - 0.8, 0.2, [{"t": stage, "size": 9, "bold": True, "color": "7F7F7F"}],
             name=f"Stufe {n}")
        text(s, x + 0.65, y + 0.30, cw3 - 0.8, 0.28, [{"t": head, "size": 13, "bold": True, "color": NAVY}],
             name=f"Titel Hebel {n}")
        text(s, x + 0.15, y + 0.66, cw3 - 0.3, 1.2, [
            {"t": body, "size": 11, "after": 3},
            {"runs": [{"t": "In Kelmis: ", "bold": True, "color": BLUE}, {"t": kelmis}], "size": 11, "after": 2},
            {"t": ref, "size": 10, "color": "7F7F7F"}], name=f"Text Hebel {n}")
        chips(s, x + 0.15, y + ch - 0.32, who)
    return s


lever_slide("Förderhebel (1/2): Grundstück, Infrastruktur und Gebäude",
            "Wo öffentliche Mittel die Baulandentwicklung und die ersten Gebäude finanzierbar machen",
            "Die Hebel 1 bis 4 setzen bei der Entwicklungsgesellschaft an und entsprechen den Maßnahmen, mit denen die "
            "Musterrechnung vom Basisszenario zum Zielbild kommt. Hebel 5 und 6 machen die Gebäude für Investoren "
            "finanzierbar. Wichtig: Im geförderten SPI-Gewerbegebiet müssen Flächen mindestens zum Verkehrswert abgegeben "
            "werden; Vergünstigungen laufen dort über Mieten oder Betrieb. Alle Angaben sind im Projekt zu prüfen.",
            levers_a)
s = lever_slide("Förderhebel (2/2): Nutzer, Betrieb und Gründung",
                "Wo Zuschüsse und Partnerschaften den Betrieb des EMRIQ-Hauses und die Gründungen tragen",
                "Die Hebel 7 bis 12 setzen beim Betrieb des EMRIQ-Hauses und bei den Gründungen an. Die Mietgarantie "
                "(Hebel 8) ist der Schlüssel, damit ein Investor das EMRIQ-Haus finanziert. Beihilferecht mitdenken: "
                "vergünstigter Kanon, Bürgschaften und Mietzuschüsse sind in der Regel Beihilfen; übliche Wege sind ein "
                "Marktwertgutachten, De-minimis (bis 300.000 € je Unternehmen in drei Jahren) oder die AGVO.",
                levers_b)
FOOT_NOTE = ("Beihilferecht: Vergünstigter Kanon, Bürgschaften und Mietzuschüsse sind in der Regel Beihilfen – "
             "Marktwertgutachten, De-minimis (300.000 € in drei Jahren) oder AGVO-Freistellung. Angaben im Projekt zu prüfen.")
for sl in (prs.slides[-2], prs.slides[-1]):
    text(sl, 0.57, 6.64, 12.18, 0.22, [{"t": FOOT_NOTE, "size": 10, "color": "7F7F7F"}], name="Hinweis Beihilferecht")

# Folie 12: Strukturierungsvorschlag aus Vorbildern
s = new_slide("Strukturierung: Was sich aus Vorbildern übertragen lässt",
              "Fünf wiederkehrende Muster internationaler Innovationsstandorte und ihr Ansatz in Kelmis",
              "Die Muster stammen aus dem Vergleich von Zürich, Cornell Tech, Cortex St. Louis, Block71, RWTH Campus, "
              "Station F, den israelischen Inkubatoren und Eindhoven. Für Kelmis sind sie bewusst klein skaliert: Es geht "
              "um ein Quartier mit einem Gründerzentrum, nicht um einen Großcampus. Die Struktur wird in AP 6 und AP 7 "
              "ausgearbeitet und mit der DG abgestimmt.")
text(s, 0.57, 1.80, 5.6, 0.3, [{"t": "Muster und Vorbilder", "size": 14, "bold": True, "color": NAVY}], name="Kopf Muster")
text(s, 7.05, 1.80, 5.6, 0.3, [{"t": "Ansatz für das EMRIQ Kelmis", "size": 14, "bold": True, "color": NAVY}],
     name="Kopf Kelmis")
patterns = [
    ("Die öffentliche Hand bleibt Eigentümerin und vergibt Baurecht oder Langzeitpacht",
     "Zürich, Cornell Tech", "Baufelder in Erbpacht (15–99 Jahre) mit Nutzungsbindung und Heimfall; Wohnbaufelder bei Bedarf im Verkauf",
     "Hebel 1"),
    ("Infrastruktur wird vorfinanziert und über Rückflüsse refinanziert",
     "Zürich, Cortex St. Louis", "Revolvierender Bodenfonds; Erschließung des Gewerbeteils mit SPI-Förderung",
     "Hebel 2–4"),
    ("Immobilie und Betrieb sind getrennt",
     "Zürich, Station F, RWTH Campus", "Investoren bauen, die WFG betreibt das EMRIQ-Haus per Generalmietvertrag",
     "Hebel 5, 8, 9"),
    ("Gründerflächen sind befristet und an das Firmenalter gekoppelt",
     "Block71, Gründerzentren", "Verträge über drei bis fünf Jahre; Anschlussflächen auf den Gewerbebaufeldern",
     "Hebel 12"),
    ("Unternehmen zahlen für Zugang",
     "RWTH Campus, Station F, Eindhoven", "Partnerschaften mit regionalen Unternehmen, Programme, Ankermieten",
     "Hebel 5, 10"),
]
for k, (pat, ex, kel, lev) in enumerate(patterns):
    y = 2.25 + k * 0.86
    box(s, 0.57, y, 12.18, 0.76, GREY1 if k % 2 == 0 else WHITE, name=f"Zeile Muster {k + 1}")
    num(s, 0.72, y + 0.19, k + 1, color=NAVY)
    text(s, 1.25, y + 0.08, 5.5, 0.62, [{"t": pat, "size": 13, "bold": True, "color": NAVY, "after": 1},
                                        {"t": "Vorbilder: " + ex, "size": 11, "color": "7F7F7F"}],
         anchor="m", name=f"Muster {k + 1}")
    text(s, 7.05, y + 0.08, 4.4, 0.62, [{"t": kel, "size": 12}], anchor="m", name=f"Kelmis {k + 1}")
    text(s, 11.55, y + 0.08, 1.1, 0.62, [{"t": lev, "size": 11, "bold": True, "color": BLUE}], anchor="m",
         align="r", name=f"Hebelbezug {k + 1}")
text(s, 0.57, 6.42, 12.18, 0.25, [{"t": "Quellen der Vorbilder: öffentliche Angaben der Standorte; einige Zahlen sind "
                                        "Plan- oder Eröffnungswerte. Übertragbarkeit auf Belgien im Projekt prüfen.",
                                   "size": 10, "color": "7F7F7F"}], name="Quellenhinweis Vorbilder")

# ---------------------------------------------------------------- Folie 8: AP 7–8
s = new_slide("AP 7–8: Umsetzung, Kommunikation und Vermarktung",
              "Priorisierte Roadmap und schlanke, verbindliche Abstimmungsformate",
              "AP 7 übersetzt die Ergebnisse in einen Umsetzungsplan mit zehn priorisierten Maßnahmen, Zuständigkeiten "
              "und Risiken. AP 8 begleitet die gesamte Bearbeitung mit festen Formaten. Akteure vor Ort und in der "
              "Euregio binden wir nur nach Freigabe durch die DG ein. Die Vermarktungsunterlagen entstehen nach AP 6 "
              "und AP 7 und werden fachlich, nicht nur sprachlich, ins Französische und Englische übertragen.")
text(s, 0.57, 1.78, 5.9, 0.3, [{"runs": [{"t": "AP 7  ", "bold": True, "color": BLUE},
                                         {"t": "Umsetzungsstrategie und Roadmap", "bold": True, "color": NAVY}],
                                "size": 15}], name="Überschrift AP 7")
rows = [("massnahmen", "Top-10-Maßnahmenkatalog mit Prioritäten, Abhängigkeiten und Entscheidungsbedarfen"),
        ("termine", "Termin- und Meilensteinplanung mit Abschluss deutlich vor dem 31.10.2028"),
        ("verantwortung", "Zuständigkeiten, Ressourcen und Schnittstellen"),
        ("verfahren", "Genehmigungs-, Planungs- und Vergabefahrplan auf Verfahrensebene"),
        ("risiko", "Risikoanalyse mit Gegenmaßnahmen")]
for i, (ic, t) in enumerate(rows):
    y = 2.25 + i * 0.86
    icon(s, ic, 0.57, y, 0.62)
    text(s, 1.40, y, 4.95, 0.62, [t], anchor="m", name=f"AP 7 Punkt {i + 1}")
text(s, 6.85, 1.78, 5.9, 0.3, [{"runs": [{"t": "AP 8  ", "bold": True, "color": BLUE},
                                         {"t": "Kommunikation, Beteiligung und Vermarktung", "bold": True,
                                          "color": NAVY}], "size": 15}], name="Überschrift AP 8")
stats = [("1", "Auftakt", "digital, bis 90 Minuten"), ("4", "Abstimmungen", "digital, je bis 45 Minuten"),
         ("2", "Workshops", "je bis 3 Stunden, einer in Präsenz"), ("1", "Abschluss", "Präsentation, digital, bis 90 Minuten")]
for i, (num, head, sub) in enumerate(stats):
    x = 6.85 + (i % 2) * 3.0
    y = 2.25 + (i // 2) * 1.32
    box(s, x, y, 2.8, 1.15, GREY1, name=f"Format {head}")
    text(s, x + 0.15, y + 0.08, 0.75, 1.0, [{"t": num, "size": 40, "bold": True, "color": NAVY}], anchor="m",
         name=f"Zahl {head}", align="c")
    text(s, x + 0.95, y + 0.1, 1.8, 0.95, [{"t": head, "size": 14, "bold": True, "color": NAVY, "after": 2},
                                           {"t": sub, "size": 13}], anchor="m", name=f"Text {head}")
box(s, 6.85, 4.98, 5.8, 1.62, NAVY, name="Vermarktung")
icon(s, "sprache", 7.07, 5.18, 0.6)
icon(s, "karte", 7.07, 5.88, 0.6)
text(s, 7.85, 5.18, 4.65, 0.6, [{"t": "Standortstory und Vermarktungsbausteine in Deutsch, Französisch und Englisch",
                                 "color": WHITE}], anchor="m", name="Text Sprachen")
text(s, 7.85, 5.88, 4.65, 0.6, [{"t": "GIS-fähige Karten im belgischen Lambert-System, kurzfristige "
                                      "Aktivierungsmaßnahmen", "color": WHITE}], anchor="m", name="Text GIS")

# ---------------------------------------------------------------- Folie 9: Zeitplan
s = new_slide("Zeitplan und Meilensteine",
              "Abschluss nach rund 16 Monaten – deutlich vor dem Endtermin am 31. Oktober 2028",
              "Sieben Meilensteine mit Freigabe durch die DG. Der Terminplan wird im Kick-off konkretisiert. Bei "
              "Beauftragung um den Jahreswechsel liegt der Abschluss im Frühjahr 2028, also mit deutlichem Puffer vor "
              "dem Endtermin. Die beiden Workshops liegen an den Stellen, an denen Entscheidungen der DG gebraucht "
              "werden: Themenprofil und Entwicklungskonzept sowie Wirtschaftlichkeit und Umsetzung.")
add_svg(s, os.path.join(SVG, "09_terminplan.svg"), 0.57, 1.80, 12.177, 4.75, "Terminplan")

# ---------------------------------------------------------------- Folie 10: Liefergegenstände
s = new_slide("Liefergegenstände und nächste Schritte",
              "Was die Deutschsprachige Gemeinschaft erhält und wie wir starten",
              "Alle Unterlagen und Rechte gehen an die DG über. Nicht enthalten sind Planungen über die Konzeptebene "
              "hinaus, Sondergutachten (z. B. Verkehr, Schall, Altlasten, Wertermittlung) und Rechtsberatung; ein "
              "solcher Bedarf wird im Umsetzungsplan mit Zeitpunkt und Kostenansatz ausgewiesen. Zum Abschluss die "
              "Fragen an das Ministerium: Betrachtungsraum, Themenpräferenz, gewünschte Formate der Ergebnisse.")
text(s, 0.57, 1.78, 7.6, 0.3, [{"t": "Liefergegenstände", "size": 15, "bold": True, "color": NAVY}],
     name="Überschrift Liefergegenstände")
items = ["Projekt- und Bearbeitungsplan mit Termin-, Risiko- und Entscheidungslogik",
         "Standort-, Markt- und Potenzialanalyse mit Standortsteckbrief",
         "Themen-, Zielgruppen- und Ecosystem-Konzept",
         "Entwicklungs- und Nutzungskonzept mit Vorzugsvariante",
         "Nachhaltigkeits-, Mobilitäts- und Infrastrukturkonzept",
         "Wirtschaftlichkeitsbetrachtung mit offengelegten Annahmen",
         "Finanzierungs-, Förder-, Träger- und Betreibermodell",
         "Umsetzungsplan mit Maßnahmen, Prioritäten und Zuständigkeiten",
         "Vermarktungsunterlagen in Deutsch, Französisch und Englisch",
         "GIS-Karten, Protokolle, Abschlussbericht mit bearbeitbaren Dateien"]
for i, t in enumerate(items):
    col, row = i // 5, i % 5
    x, y = 0.57 + col * 3.9, 2.22 + row * 0.62
    add_svg(s, os.path.join(SVG, "icons", "haken.svg"), x, y + 0.05, 0.26, 0.26, f"Haken L{i + 1}",
            color_swap={"#002965": "#" + GREEN})
    text(s, x + 0.38, y, 3.35, 0.56, [{"t": t, "size": 13}], name=f"Liefergegenstand {i + 1}")
box(s, 0.57, 5.40, 7.55, 1.2, GREY1, name="Qualitätssicherung")
icon(s, "qualitaet", 0.77, 5.62, 0.75)
text(s, 1.72, 5.48, 6.25, 1.05, [{"t": "Qualitätssicherung", "size": 14, "bold": True, "color": NAVY, "after": 2},
                                 {"t": "Vier-Augen-Prinzip, fachübergreifende Reviews an den Meilensteinen, "
                                       "dokumentierte Freigaben, fachliche Prüfung der französischen und englischen "
                                       "Fassung", "size": 13}], anchor="m", name="Text Qualitätssicherung")
box(s, 8.45, 1.80, 4.3, 4.80, NAVY, name="Nächste Schritte")
text(s, 8.70, 2.00, 3.85, 0.4, [{"t": "Nächste Schritte", "size": 16, "bold": True, "color": WHITE}],
     name="Überschrift Nächste Schritte")
steps = ["Offene Punkte aus dem heutigen Termin klären: Betrachtungsraum, Themenpräferenz, Formate der Ergebnisse",
         "Angebot auf dieser Grundlage abschließen",
         "Beauftragung und Auftakt im ersten Monat",
         "Bereitstellung der Grundlagen: Fachgutachten, Gewerbeflächenstrategie, MORO-Bericht, Geodaten"]
yy = 2.55
for i, t in enumerate(steps):
    n = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.70), Inches(yy), Inches(0.42), Inches(0.42))
    n.name = f"Nummer {i + 1}"
    fill(n, BLUE)
    tf = n.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(i + 1)
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = rgb(WHITE)
    h = 0.95 if len(t) > 60 else 0.62
    text(s, 9.30, yy - 0.02, 3.25, h, [{"t": t, "color": WHITE}], name=f"Schritt {i + 1}")
    yy += h + 0.12

prs.core_properties.title = "EMRIQ Kelmis – Leistungskonzept"
prs.core_properties.author = ""
prs.core_properties.last_modified_by = ""
os.makedirs(os.path.dirname(OUT), exist_ok=True)
prs.save(OUT)
print("geschrieben:", OUT, len(prs.slides._sldIdLst), "Folien")

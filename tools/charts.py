"""Grafiken für den Word-Bericht EMRIQ Kelmis (statisch, Druck/Lichtmodus).
Alle Modellparameter stehen in PARAMS und werden auch für die Tabellen im Bericht verwendet
(Ausgabe: bericht/modell_ergebnisse.json)."""
import json, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

OUT = os.path.join(os.path.dirname(__file__), "..", "bericht", "grafiken")
os.makedirs(OUT, exist_ok=True)

# ---- Design-Tokens (Referenzpalette dataviz, Lichtmodus) ----
INK, INK2, MUTED, GRID = "#0b0b0b", "#52514e", "#8a8984", "#e4e3df"
S1, S2, S3 = "#2a78d6", "#eb6834", "#1baf7a"  # Kategorial Slot 1-3
BLUE_250, BLUE_100, NAVY = "#86b6ef", "#cde2fb", "#0d366b"
NEUTRAL = "#b9b8b2"
for f in font_manager.findSystemFonts():
    if "Inter" in f:
        font_manager.fontManager.addfont(f)
plt.rcParams.update({
    "font.family": "Inter", "font.size": 9, "text.color": INK, "axes.labelcolor": INK2,
    "xtick.color": INK2, "ytick.color": INK2, "axes.edgecolor": GRID, "axes.linewidth": 0.8,
    "figure.dpi": 220, "savefig.bbox": "tight", "savefig.pad_inches": 0.08,
})

def de(v, nd=1, sign=False):
    t = f"{v:+.{nd}f}" if sign else f"{v:.{nd}f}"
    return t.replace(".", ",")

from matplotlib.ticker import FuncFormatter
DEFMT = FuncFormatter(lambda v, _: (f"{v:.2f}" if abs(v - round(v)) > 1e-9 and abs(v*2 - round(v*2)) > 1e-9 else (f"{v:.1f}" if abs(v - round(v)) > 1e-9 else f"{v:.0f}")).replace(".", ",").replace("-", "−"))

def style(ax, grid_axis="x"):
    ax.xaxis.set_major_formatter(DEFMT) if grid_axis == "x" else ax.yaxis.set_major_formatter(DEFMT)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.grid(axis=grid_axis, color=GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(length=0)

# ================= Modell Baulandentwicklung (Annahmen, Mio. €) =================
PARAMS = {
    "brutto_ha": 8.0, "anteil_oeffentlich": 0.25, "anteil_wohnen": 0.55,
    # Grundstückswertanteile €/m² BGF kalibriert nach 06b (Wohnen 60–150, Gewerbe/Labor 30–80; Zielbild mit Quartierseffekt)
    "szenarien": {
        "Konservativ": {"gfz": 0.75, "gw_wohnen": 60,  "gw_gewerbe": 30, "grunderwerb": 4.0, "erschliessung": 10.0, "foerderquote": 0.40, "eigentuemer": 0.5},
        "Basis":       {"gfz": 1.00, "gw_wohnen": 110, "gw_gewerbe": 55, "grunderwerb": 4.0, "erschliessung": 10.0, "foerderquote": 0.50, "eigentuemer": 1.5},
        "Zielbild":    {"gfz": 1.25, "gw_wohnen": 180, "gw_gewerbe": 80, "grunderwerb": 2.5, "erschliessung": 8.5,  "foerderquote": 0.70, "eigentuemer": 2.5},
    },
    "kosten_fix": {"Verlagerung Busbetriebshof": 2.5, "Abriss / Altlasten": 3.0,
                    "Planung, Gutachten, Verfahren": 1.3, "Vermarktung, Management, Finanzierung": 1.5},
}

def rechne(name):
    p = PARAMS["szenarien"][name]
    netto_m2 = PARAMS["brutto_ha"] * 10000 * (1 - PARAMS["anteil_oeffentlich"])
    bgf = netto_m2 * p["gfz"]
    bgf_w, bgf_g = bgf * PARAMS["anteil_wohnen"], bgf * (1 - PARAMS["anteil_wohnen"])
    erl_w, erl_g = bgf_w * p["gw_wohnen"] / 1e6, bgf_g * p["gw_gewerbe"] / 1e6
    kf = PARAMS["kosten_fix"]
    kosten = p["grunderwerb"] + p["erschliessung"] + sum(kf.values())
    foerder = p["foerderquote"] * (p["erschliessung"] + kf["Abriss / Altlasten"])
    einnahmen = erl_w + erl_g + foerder + p["eigentuemer"]
    return {"bgf": bgf, "bgf_w": bgf_w, "bgf_g": bgf_g, "erl_w": erl_w, "erl_g": erl_g, "baufelder": erl_w + erl_g,
            "foerder": foerder, "eigentuemer": p["eigentuemer"], "kosten": kosten, "einnahmen": einnahmen,
            "saldo": einnahmen - kosten, "deckung": einnahmen / kosten, **p}

ERG = {n: rechne(n) for n in PARAMS["szenarien"]}

# Zeitliche Verteilung (Anteile je Jahr) für den kumulierten Cashflow
JAHRE = list(range(2027, 2037))
KOSTEN_T = [0.08, 0.14, 0.26, 0.17, 0.12, 0.12, 0.10, 0.01, 0.0, 0.0]
FOERDER_T = [0, 0, 0.31, 0.31, 0, 0.20, 0.18, 0, 0, 0]
EIGENT_T = [0, 0, 0.33, 0, 0, 0.67, 0, 0, 0, 0]
BAUF_T = [0, 0, 0.11, 0.14, 0.11, 0.16, 0.16, 0.115, 0.115, 0.09]

def cashflow(e):
    cum, out = 0.0, []
    for k in range(len(JAHRE)):
        cum += -e["kosten"] * KOSTEN_T[k] + e["foerder"] * FOERDER_T[k] + e["eigentuemer"] * EIGENT_T[k] + e["baufelder"] * BAUF_T[k]
        out.append(cum)
    return out

CF = {n: cashflow(e) for n, e in ERG.items()}
for n in ERG:
    ERG[n]["spitzenbedarf"] = min(CF[n])
    ERG[n]["spitzenjahr"] = JAHRE[CF[n].index(min(CF[n]))]

# ================= 1 Cluster-Scoring =================
def chart_cluster():
    data = [
        ("Smart Energy / Netze", 4.35, "Enabler"),
        ("Einstein-Teleskop & Präzisionstechnik", 4.25, "Themenbildend"),
        ("Zirkuläres Bauen & Holz", 4.25, "Themenbildend"),
        ("Digital / KI / Software", 4.05, "Enabler"),
        ("Space & Robotik / Automation", 3.65, "Anwendungsmarkt"),
        ("Sicherheit / Dual-Use (mit Negativliste)", 3.55, "Anwendungsmarkt"),
        ("Autonome Mobilität", 3.55, "Enabler"),
        ("Kultur / Tourismus", 3.55, "Enabler"),
        ("Health / MedTech", 3.50, "Enabler"),
        ("Wasserstoff", 3.20, "Enabler"),
        ("Advanced Materials / Kreislauf-Metalle", 3.15, "Enabler"),
    ][::-1]
    col = {"Themenbildend": S1, "Anwendungsmarkt": BLUE_250, "Enabler": NEUTRAL}
    fig, ax = plt.subplots(figsize=(7.2, 4.1))
    y = range(len(data))
    ax.barh(y, [d[1] for d in data], color=[col[d[2]] for d in data], height=0.62, edgecolor="white", linewidth=1)
    for k, d in enumerate(data):
        ax.text(d[1] + 0.03, k, f"{d[1]:.2f}".replace(".", ","), va="center", fontsize=8, color=INK)
    ax.set_yticks(list(y)); ax.set_yticklabels([d[0] for d in data], fontsize=8.5, color=INK)
    ax.set_xlim(2.5, 4.7); ax.set_xlabel("Gewichteter Score (1–5)")
    style(ax)
    handles = [plt.Rectangle((0, 0), 1, 1, color=col[k]) for k in ("Themenbildend", "Anwendungsmarkt", "Enabler")]
    ax.legend(handles, ["Themenbildend (Hauptthema-Kandidat)", "Anwendungsmarkt des Hauptthemas", "Enabler / Querschnitt / Nische"],
              loc="lower right", frameon=False, fontsize=8)
    fig.savefig(os.path.join(OUT, "01_cluster_scoring.png")); plt.close(fig)

# ================= 2 Wertschöpfungskurve & Sweet Spot =================
def chart_wert():
    stufen = ["Brache /\nBestand", "Planungsrecht\ngesichert", "Erschließung im Bau\n+ Bodenattest", "Leuchtturm in Betrieb\n+ Ankermieter", "Quartier\netabliert"]
    mitte = [100, 145, 195, 250, 300]
    lo = [100, 130, 170, 220, 260]
    hi = [100, 160, 220, 280, 350]
    risiko = [100, 70, 45, 25, 15]
    x = range(len(stufen))
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    ax.axvspan(1.65, 3.35, color=BLUE_100, alpha=0.55, lw=0)
    ax.text(2.5, 345, "Sweet Spot für Baufeldverkäufe", ha="center", fontsize=8.5, color=NAVY, fontweight="bold")
    ax.fill_between(x, lo, hi, color=S1, alpha=0.12, lw=0)
    ax.plot(x, mitte, color=S1, lw=2, marker="o", ms=6, label="Bodenwert-Index (Bandbreite schattiert)")
    ax.plot(x, risiko, color=S2, lw=2, marker="o", ms=6, label="Entwicklungsrisiko-Index")
    for k, v in enumerate(mitte):
        ax.text(k, v + 14, str(v), ha="center", fontsize=8, color=INK)
    ax.set_xticks(list(x)); ax.set_xticklabels(stufen, fontsize=7.8, color=INK)
    ax.set_ylim(0, 370); ax.set_ylabel("Index (Brache = 100)")
    style(ax, "y")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=2, frameon=False, fontsize=8)
    fig.savefig(os.path.join(OUT, "02_wertschoepfung_sweetspot.png")); plt.close(fig)

# ================= 3 Refinanzierung Szenarien =================
def chart_refi():
    names = list(ERG)[::-1]
    fig, ax = plt.subplots(figsize=(7.2, 2.9))
    y = range(len(names))
    left = [0] * len(names)
    parts = [("baufelder", "Baufelderlöse / kap. Erbpacht", S1), ("foerder", "Fördermittel Erschließung", S2), ("eigentuemer", "Beiträge privater Eigentümer", S3)]
    for key, lab, c in parts:
        vals = [ERG[n][key] for n in names]
        ax.barh(y, vals, left=left, color=c, height=0.55, label=lab, edgecolor="white", linewidth=1.5)
        left = [a + b for a, b in zip(left, vals)]
    for k, n in enumerate(names):
        e = ERG[n]
        ax.plot([e["kosten"], e["kosten"]], [k - 0.36, k + 0.36], color=INK, lw=2)
        ax.text(max(e["kosten"], e["einnahmen"]) + 0.5, k, f"Deckung {e['deckung']*100:.0f} %  ·  Saldo {de(e['saldo'], 1, True)} Mio. €",
                va="center", fontsize=8, color=INK)
    ax.plot([], [], color=INK, lw=2, label="Kosten Baulandentwicklung")
    ax.set_yticks(list(y)); ax.set_yticklabels(names, fontsize=9, color=INK)
    ax.set_xlim(0, 36); ax.set_xlabel("Mio. € (nominal, ohne MwSt.)")
    style(ax)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.28), ncol=4, frameon=False, fontsize=7.5)
    fig.savefig(os.path.join(OUT, "03_refinanzierung_szenarien.png")); plt.close(fig)

# ================= 4 Kumulierter Cashflow =================
def chart_cf():
    fig, ax = plt.subplots(figsize=(7.2, 3.3))
    cols = {"Konservativ": S2, "Basis": S1, "Zielbild": S3}
    ax.axhline(0, color=INK2, lw=0.9)
    for n, vals in CF.items():
        ax.plot(JAHRE, vals, color=cols[n], lw=2, marker="o", ms=4, label=n)
        e = ERG[n]
        ax.text(JAHRE[-1] + 0.15, vals[-1], f"{n}: {de(vals[-1], 1, True)}", va="center", fontsize=8, color=INK)
    b = ERG["Basis"]
    ax.annotate(f"Spitzenkapitalbedarf Basis\n≈ {de(b['spitzenbedarf'])} Mio. € ({b['spitzenjahr']})",
                xy=(b["spitzenjahr"], b["spitzenbedarf"]), xytext=(2026.9, -16.6),
                fontsize=8, color=INK, arrowprops=dict(arrowstyle="-", color=INK2, lw=0.8))
    ax.set_xlim(2026.6, 2038.6); ax.set_xticks(JAHRE)
    ax.set_ylabel("kumuliert, Mio. €")
    style(ax, "y")
    ax.set_ylim(-18, 4.5)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, 1.13), ncol=3, frameon=False, fontsize=8)
    fig.savefig(os.path.join(OUT, "04_cashflow_bodenfonds.png")); plt.close(fig)

# ================= 5 Kapitalkreislauf (Schema) =================
def chart_kreislauf():
    import math
    fig, ax = plt.subplots(figsize=(7.2, 4.6))
    ax.set_xlim(0, 100); ax.set_ylim(0, 64); ax.axis("off"); ax.set_aspect("auto")
    steps = [
        ("1  Startkapital", "DG · Gemeinde/AGR · EU\nrevolvierender Bodenfonds", NAVY, "white"),
        ("2  De-Risking", "Baurecht · Bodenattest\nErschließung", S1, "white"),
        ("3  Leuchtturm bauen", "EMRIQ-Haus · Boarding\n„Skin in the Game“", S2, "white"),
        ("4  Baufelder verkaufen", "im Sweet Spot · Konzept-\nvergabe · kap. Erbpacht", S1, "white"),
        ("5  Markt baut Quartier", "Entwickler · Investoren\nEigennutzer", "#e9eef5", INK),
        ("6  Kapitalrecycling", "Erlöse + Verkauf Leuchtturm\nnach Stabilisierung", S3, "white"),
    ]
    cx, cy, rx, ry = 50, 31, 36, 21
    w, h = 27, 12.5
    pos = []
    for k in range(6):
        ang = math.radians(90 - k * 60 + 30)
        pos.append((cx + rx * math.cos(ang), cy + ry * math.sin(ang)))
    # Reihenfolge im Uhrzeigersinn ab oben links
    order = [(50 - 30, 54), (50 + 30, 54), (50 + 38, 31), (50 + 30, 8), (50 - 30, 8), (50 - 38, 31)]
    order = [(20, 53), (80, 53), (86, 31), (80, 9), (20, 9), (14, 31)]
    for (x, y), (t, sub, fill, tc) in zip(order, steps):
        ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.3,rounding_size=1.8", fc=fill, ec="none"))
        ax.text(x, y + 2.6, t, ha="center", va="center", fontsize=8.6, fontweight="bold", color=tc)
        ax.text(x, y - 2.3, sub, ha="center", va="center", fontsize=6.9, color=tc, linespacing=1.25)
    def arr(a, b, rad):
        ax.add_patch(FancyArrowPatch(a, b, arrowstyle="-|>", mutation_scale=11, color=INK2, lw=1.1, connectionstyle=f"arc3,rad={rad}"))
    arr((34.5, 53), (65.5, 53), 0.0)
    arr((89, 46), (89.5, 38), -0.3)
    arr((89.5, 24), (89, 16), -0.3)
    arr((65.5, 9), (34.5, 9), 0.0)
    arr((11, 16), (10.5, 24), -0.3)
    arr((10.5, 38), (11, 46), -0.3)
    ax.text(50, 33, "Revolvierender\nKapitalkreislauf", ha="center", va="center", fontsize=11, fontweight="bold", color=NAVY, linespacing=1.2)
    ax.text(50, 25.5, "Der Initiator entwickelt Boden und Vertrauen –\nder Markt baut das Quartier.", ha="center", va="center", fontsize=7.6, color=INK2, linespacing=1.3)
    fig.savefig(os.path.join(OUT, "05_kapitalkreislauf.png")); plt.close(fig)

# ================= 6 Brücke Basis -> Zielbild =================
def bruecke():
    b, z = ERG["Basis"], ERG["Zielbild"]
    p_b, p_z = PARAMS["szenarien"]["Basis"], PARAMS["szenarien"]["Zielbild"]
    hebel = [
        ("Bodenwert\n& Dichte", z["baufelder"] - b["baufelder"]),
        ("Förderung\nGewerbeteil\n(SPI-Operator)", z["foerder"] - b["foerder"]),
        ("Kooperative\nBauland-\nentwicklung", (p_b["grunderwerb"] - p_z["grunderwerb"]) + (z["eigentuemer"] - b["eigentuemer"])),
        ("Schlanke\nErschließung", p_b["erschliessung"] - p_z["erschliessung"]),
    ]
    return b["saldo"], hebel, z["saldo"]

def chart_bruecke():
    start, hebel, ende = bruecke()
    labels = ["Saldo\nBasis"] + [h[0] for h in hebel] + ["Saldo\nZielbild"]
    fig, ax = plt.subplots(figsize=(7.2, 3.5))
    ax.axhline(0, color=INK2, lw=0.9)
    ax.bar(0, start, color=S2, width=0.6)
    ax.text(0, start - 0.6, f"{de(start, 1, True)}", ha="center", va="top", fontsize=8.5, color=INK, fontweight="bold")
    cum = start
    for k, (lab, v) in enumerate(hebel, start=1):
        ax.bar(k, v, bottom=cum, color=S1, width=0.6, edgecolor="white", linewidth=1.5)
        ax.text(k, cum + v + 0.35, f"+{de(round(v + 1e-9, 1))}", ha="center", va="bottom", fontsize=8.5, color=INK)
        ax.plot([k - 1 + 0.3, k - 0.3], [cum, cum], color=MUTED, lw=0.8, ls=(0, (2, 2)))
        cum += v
    k = len(hebel) + 1
    ax.plot([k - 1 + 0.3, k - 0.3], [cum, cum], color=MUTED, lw=0.8, ls=(0, (2, 2)))
    ax.bar(k, ende, color=S3, width=0.6)
    ax.text(k, ende + 0.35, f"{de(ende, 1, True)}", ha="center", va="bottom", fontsize=8.5, color=INK, fontweight="bold")
    ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels, fontsize=7.6, color=INK)
    ax.set_ylabel("Saldo Baulandentwicklung, Mio. €")
    ax.set_ylim(min(start, 0) - 2.5, max(ende, 0) + 3)
    style(ax, "y")
    fig.savefig(os.path.join(OUT, "06_bruecke_refinanzierung.png")); plt.close(fig)

if __name__ == "__main__":
    chart_cluster(); chart_wert(); chart_refi(); chart_cf(); chart_kreislauf(); chart_bruecke()
    _s, _h, _e = bruecke(); print("Brücke:", round(_s,2), [(h[0].replace(chr(10)," "), round(h[1],2)) for h in _h], round(_e,2))
    res = {n: {k: (round(v, 2) if isinstance(v, float) else v) for k, v in e.items()} for n, e in ERG.items()}
    res["_cashflow"] = {n: [round(v, 2) for v in vals] for n, vals in CF.items()}
    json.dump(res, open(os.path.join(OUT, "..", "modell_ergebnisse.json"), "w"), ensure_ascii=False, indent=1)
    for n, e in ERG.items():
        print(f"{n:13s} BGF {e['bgf']:7.0f} m²  Baufelder {e['baufelder']:5.2f}  Förder {e['foerder']:4.2f}  Eigent {e['eigentuemer']:3.1f}  "
              f"Kosten {e['kosten']:5.2f}  Saldo {e['saldo']:+5.2f}  Deckung {e['deckung']*100:4.0f}%  Spitze {e['spitzenbedarf']:5.2f} ({e['spitzenjahr']})")

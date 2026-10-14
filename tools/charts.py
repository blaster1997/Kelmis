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
# Nüchterne Office-Farben
INK, INK2, MUTED, GRID = "#000000", "#404040", "#7f7f7f", "#d9d9d9"
S1, S2, S3 = "#4472c4", "#ed7d31", "#70ad47"
BLUE_250, BLUE_100, NAVY = "#9dc3e6", "#e7e6e6", "#1f3864"
NEUTRAL = "#a5a5a5"
for f in font_manager.findSystemFonts():
    if "Liberation" in f:
        font_manager.fontManager.addfont(f)
plt.rcParams.update({
    "font.family": "Liberation Sans", "font.size": 9, "text.color": INK, "axes.labelcolor": INK2,
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
        ("Energie / Netze", 4.35, "Enabler"),
        ("Einstein-Teleskop / Präzisionstechnik", 4.25, "Themenbildend"),
        ("Zirkuläres Bauen / Holz", 4.25, "Themenbildend"),
        ("Digitalwirtschaft / KI", 4.05, "Enabler"),
        ("Raumfahrt / Robotik / Automation", 3.65, "Anwendungsmarkt"),
        ("Sicherheit / Dual-Use (mit Ausschlüssen)", 3.55, "Anwendungsmarkt"),
        ("Autonome Mobilität", 3.55, "Enabler"),
        ("Kultur / Tourismus", 3.55, "Enabler"),
        ("Gesundheit / Medizintechnik", 3.50, "Enabler"),
        ("Wasserstoff", 3.20, "Enabler"),
        ("Werkstoffe / Kreislaufmetalle", 3.15, "Enabler"),
    ][::-1]
    col = {"Themenbildend": S1, "Anwendungsmarkt": BLUE_250, "Enabler": NEUTRAL}
    fig, ax = plt.subplots(figsize=(7.2, 4.1))
    y = range(len(data))
    ax.barh(y, [d[1] for d in data], color=[col[d[2]] for d in data], height=0.62, edgecolor="white", linewidth=1)
    for k, d in enumerate(data):
        ax.text(d[1] + 0.03, k, f"{d[1]:.2f}".replace(".", ","), va="center", fontsize=8, color=INK)
    ax.set_yticks(list(y)); ax.set_yticklabels([d[0] for d in data], fontsize=8.5, color=INK)
    ax.set_xlim(2.5, 4.7); ax.set_xlabel("Gewichtete Bewertung (1 = gering, 5 = hoch)")
    style(ax)
    handles = [plt.Rectangle((0, 0), 1, 1, color=col[k]) for k in ("Themenbildend", "Anwendungsmarkt", "Enabler")]
    ax.legend(handles, ["themenbildend", "Anwendungsfeld des Hauptthemas", "Querschnitts- oder Ergänzungsthema"],
              loc="lower right", frameon=False, fontsize=8)
    fig.savefig(os.path.join(OUT, "01_cluster_scoring.png")); plt.close(fig)

# ================= 2 Wertschöpfungskurve & Sweet Spot =================
def chart_wert():
    stufen = ["Brache,\nBestand", "Baurecht\ngesichert", "Erschließung im Bau,\nBodenzustand geklärt", "erstes Gebäude in\nBetrieb, Ankermieter", "Quartier\netabliert"]
    mitte = [100, 145, 195, 250, 300]
    lo = [100, 130, 170, 220, 260]
    hi = [100, 160, 220, 280, 350]
    risiko = [100, 70, 45, 25, 15]
    x = range(len(stufen))
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    ax.axvspan(1.65, 3.35, color="#f2f2f2", lw=0)
    ax.text(2.5, 345, "günstiger Zeitraum für die Vergabe", ha="center", fontsize=8.5, color=INK)
    ax.fill_between(x, lo, hi, color=S1, alpha=0.15, lw=0)
    ax.plot(x, mitte, color=S1, lw=1.8, marker="s", ms=5, label="Bodenwert (Index, Fläche = Bandbreite)")
    ax.plot(x, risiko, color=S2, lw=1.8, marker="s", ms=5, label="Entwicklungsrisiko (Index)")
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
    parts = [("baufelder", "Erlöse Baufelder", S1), ("foerder", "Fördermittel", S2), ("eigentuemer", "Beiträge Eigentümer", NEUTRAL)]
    for key, lab, c in parts:
        vals = [ERG[n][key] for n in names]
        ax.barh(y, vals, left=left, color=c, height=0.55, label=lab, edgecolor="white", linewidth=1.5)
        left = [a + b for a, b in zip(left, vals)]
    for k, n in enumerate(names):
        e = ERG[n]
        ax.plot([e["kosten"], e["kosten"]], [k - 0.36, k + 0.36], color=INK, lw=2)
        ax.text(max(e["kosten"], e["einnahmen"]) + 0.5, k, f"Deckung {e['deckung']*100:.0f} %, Saldo {de(e['saldo'], 1, True)} Mio. €",
                va="center", fontsize=8, color=INK)
    ax.plot([], [], color=INK, lw=2, label="Kosten")
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
        ax.plot(JAHRE, vals, color=cols[n], lw=1.8, marker="s", ms=3.5, label=n)
        e = ERG[n]
        ax.text(JAHRE[-1] + 0.15, vals[-1], f"{n}: {de(vals[-1], 1, True)}", va="center", fontsize=8, color=INK)
    b = ERG["Basis"]
    ax.annotate(f"max. Kapitalbedarf Basis:\nrd. {de(abs(b['spitzenbedarf']))} Mio. € ({b['spitzenjahr']})",
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
    from matplotlib.patches import Rectangle
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    ax.set_xlim(0, 100); ax.set_ylim(0, 60); ax.axis("off")
    steps = [
        ("1  Startkapital", "DG, Gemeinde/AGR, Fördermittel;\nrevolvierender Bodenfonds"),
        ("2  Baureife herstellen", "Baurecht, Bodensanierung,\nErschließung"),
        ("3  Pilotgebäude", "EMRIQ-Haus,\nWohnmodul"),
        ("4  Vergabe der Baufelder", "in Tranchen; Verkauf oder\nkapitalisierte Erbpacht"),
        ("5  Bebauung durch Dritte", "Bauträger, Investoren,\nEigennutzer"),
        ("6  Rückfluss", "Erlöse, Verkauf eigener\nObjekte nach Vermietung"),
    ]
    pos = [(20, 50), (80, 50), (86, 30), (80, 10), (20, 10), (14, 30)]
    w, h = 28, 12
    for (x, y), (t, sub) in zip(pos, steps):
        ax.add_patch(Rectangle((x - w / 2, y - h / 2), w, h, fc="#f2f2f2", ec=INK2, lw=0.8))
        ax.text(x, y + 2.4, t, ha="center", va="center", fontsize=8.5, fontweight="bold", color=INK)
        ax.text(x, y - 2.2, sub, ha="center", va="center", fontsize=7.2, color=INK, linespacing=1.2)
    def arr(a, b):
        ax.add_patch(FancyArrowPatch(a, b, arrowstyle="-|>", mutation_scale=10, color=INK2, lw=0.9))
    arr((34.5, 50), (65.5, 50)); arr((88, 43.8), (88, 36.2)); arr((88, 23.8), (88, 16.2))
    arr((65.5, 10), (34.5, 10)); arr((12, 16.2), (12, 23.8)); arr((12, 36.2), (12, 43.8))
    ax.text(50, 30, "Kapitalkreislauf der\nBaulandentwicklung", ha="center", va="center", fontsize=10, color=INK2, linespacing=1.3)
    fig.savefig(os.path.join(OUT, "05_kapitalkreislauf.png")); plt.close(fig)

# ================= 6 Brücke Basis -> Zielbild =================
def bruecke():
    b, z = ERG["Basis"], ERG["Zielbild"]
    p_b, p_z = PARAMS["szenarien"]["Basis"], PARAMS["szenarien"]["Zielbild"]
    hebel = [
        ("Bodenwert\nund Dichte", z["baufelder"] - b["baufelder"]),
        ("Förderung\nGewerbeteil\n(über SPI)", z["foerder"] - b["foerder"]),
        ("Kooperative\nBauland-\nentwicklung", (p_b["grunderwerb"] - p_z["grunderwerb"]) + (z["eigentuemer"] - b["eigentuemer"])),
        ("Sparsame\nErschließung", p_b["erschliessung"] - p_z["erschliessung"]),
    ]
    return b["saldo"], hebel, z["saldo"]

def chart_bruecke():
    start, hebel, ende = bruecke()
    labels = ["Saldo\nBasis"] + [h[0] for h in hebel] + ["Saldo\nZielbild"]
    fig, ax = plt.subplots(figsize=(7.2, 3.5))
    ax.axhline(0, color=INK2, lw=0.9)
    ax.bar(0, start, color=NEUTRAL, width=0.6)
    ax.text(0, start - 0.6, f"{de(start, 1, True)}", ha="center", va="top", fontsize=8.5, color=INK, fontweight="bold")
    cum = start
    for k, (lab, v) in enumerate(hebel, start=1):
        ax.bar(k, v, bottom=cum, color=S1, width=0.6, edgecolor="white", linewidth=1.5)
        ax.text(k, cum + v + 0.35, f"+{de(round(v + 1e-9, 1))}", ha="center", va="bottom", fontsize=8.5, color=INK)
        ax.plot([k - 1 + 0.3, k - 0.3], [cum, cum], color=MUTED, lw=0.8, ls=(0, (2, 2)))
        cum += v
    k = len(hebel) + 1
    ax.plot([k - 1 + 0.3, k - 0.3], [cum, cum], color=MUTED, lw=0.8, ls=(0, (2, 2)))
    ax.bar(k, ende, color=NEUTRAL, width=0.6)
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

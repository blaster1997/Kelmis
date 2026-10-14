# EMRIQ Kelmis – Briefing, Analysen & Endbericht

Analyse- und Akquiseunterlagen zum Verfahren „Städtebauliches Entwicklungskonzept mit Umsetzungsplan zur Implementierung eines Euregio Maas-Rhein Innovations-Quartiers (EMRIQ)“ (MDG, Ref. FbRAUM.KaR/39.02-01/26.128).

## Lieferdokumente (Word, gelayoutet)

| Datei | Inhalt |
|---|---|
| `export/EMRIQ_Kelmis_Endbericht.docx` (+ PDF) | **Endbericht, kundentauglich:** Management Summary, Standort-, Markt- und Branchenclusteranalyse, strategische Quartiersentwicklung (15 Hebel), Entwicklungsmodell und Kapitalmarktfähigkeit, Umsetzung, Förderung und Wirtschaftlichkeit, Risiken |
| `export/EMRIQ_Kelmis_Workshop_Briefing_intern.docx` (+ PDF) | **Intern, nicht weitergeben:** Bewertungslogik, Wettbewerb, Ablaufplan, Musterlieferobjekte, Kalkulation, Workshop-Drehbuch |

## Quellen und Teilberichte

| Datei | Inhalt |
|---|---|
| `bericht/EMRIQ_Kelmis_Endbericht.md` | Markdown-Quelle des Endberichts |
| `bericht/EMRIQ_Kelmis_Workshop_Briefing_intern.md` | Markdown-Quelle des internen Briefings (abgeleitet aus 05) |
| `bericht/grafiken/` · `bericht/modell_ergebnisse.json` | Grafiken und Ergebnisse der Musterrechnung |
| `00_Kontext/Quellenauswertung.md` | Lastenheft, Forum, Gewerbeflächenstrategie, OBL'EU, Ostbelgien 2040 |
| `01_Standortanalyse_Kelmis.md` | Makro-/Mikrostandort, Restriktionen, Suchräume, SWOT |
| `02_Marktanalyse.md` | Wohnen, Büro/Labor, Ökosystem, Nachfrage, Zielgruppen |
| `03a_Branchencluster_Hightech.md` | Einstein-Teleskop/Präzision, Dual-Use, Materials, Space/Robotik |
| `03b_Branchencluster_Digital_Bau_Health_Energie.md` | Digital, Holz/zirkulär, Health, Energie, Kultur |
| `04_Foerderung_Governance_Wirtschaftlichkeit.md` | Förderkulisse, Governance, Kostenkennwerte, Musterrechnung Phase 1 |
| `05_Akquise_und_Workshop_Strategie.md` | Bieter- und Workshopstrategie (Quelle des internen Briefings) |
| `06a_Kapitalmarkt_Investoren.md` | Renditen, Investorengruppen, Strukturen, Benchmarks Master Developer |
| `06b_Bodenwerte_Recht_Steuern.md` | Bodenwerte, Erschließungsförderung (Dekret 2017), Erbpacht, Steuern, Bodendekret |

## Neu bauen

```bash
python3 tools/charts.py   # Musterrechnung + Grafiken
bash tools/build.sh bericht/EMRIQ_Kelmis_Endbericht.md export/EMRIQ_Kelmis_Endbericht.docx "EMRIQ Kelmis – Endbericht" bericht/cover_endbericht.json
bash tools/build.sh bericht/EMRIQ_Kelmis_Workshop_Briefing_intern.md export/EMRIQ_Kelmis_Workshop_Briefing_intern.docx "EMRIQ Kelmis – Workshop-Briefing (intern)" bericht/cover_workshop.json
```

Benötigt: Node mit `docx`, Python mit `matplotlib`, LibreOffice (für PDF und Seitenzahlen im Inhaltsverzeichnis).

Stand: 14.10.2026 · intern/vertraulich

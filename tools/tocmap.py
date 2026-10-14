"""Ermittelt die Seitenzahlen von Überschriften und Beschriftungen aus dem PDF des ersten Durchlaufs."""
import json, subprocess, sys, re
pdf, marks_file, out = sys.argv[1:4]
marks = json.load(open(marks_file))
n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
squash = lambda s: re.sub(r"[\s\-­]+", "", s)
pages = []
for p in range(1, n + 1):
    t = subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), pdf, "-"], capture_output=True, text=True).stdout
    pages.append((squash(t), "....." in t))  # Verzeichnisseiten erkennt man an Punktführern
DIRS = {"Abbildungsverzeichnis", "Tabellenverzeichnis"}
result, start = {}, 1
for mk in marks:
    key = squash(mk)[:45]
    order = range(n, 0, -1) if mk in DIRS else range(start, n + 1)  # Verzeichnisse stehen am Ende
    for p in order:
        txt, is_dir = pages[p - 1]
        if key in txt and (not is_dir or mk in DIRS):
            result[mk] = p
            if not mk.startswith(("Abbildung ", "Tabelle ")):
                start = p
            break
json.dump(result, open(out, "w"), ensure_ascii=False)
print(f"{len(result)}/{len(marks)} Marken verortet")

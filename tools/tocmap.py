"""Ermittelt die Seitenzahl jeder Überschrift aus dem PDF des ersten Durchlaufs."""
import json, subprocess, sys, re
pdf, headings_file, out = sys.argv[1:4]
headings = json.load(open(headings_file))
n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
pages = [subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), "-layout", pdf, "-"], capture_output=True, text=True).stdout for p in range(1, n + 1)]
norm = lambda s: re.sub(r"\s+", " ", s).strip()
npages = [norm(p) for p in pages]
result, start = {}, 1
for h in headings:
    key = norm(h)[:40]
    for p in range(start, n + 1):
        # Inhaltsverzeichnis-Seiten überspringen (dort stehen Punktführer)
        if key in npages[p - 1] and "....." not in npages[p - 1][max(0, npages[p - 1].find(key)):npages[p - 1].find(key) + len(key) + 12]:
            result[h] = p; start = p; break
json.dump(result, open(out, "w"), ensure_ascii=False)
print(f"{len(result)}/{len(headings)} Überschriften verortet")

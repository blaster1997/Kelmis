#!/usr/bin/env bash
# 2-Pass-Build: Markdown -> DOCX (Pass 1) -> PDF -> Seitenzahlen der Überschriften -> DOCX mit Inhaltsverzeichnis (Pass 2) -> PDF
# Aufruf: tools/build.sh <input.md> <output.docx> "<Kopfzeile>" <cover.json>
set -euo pipefail
IN="$(realpath "$1")"; OUT="$(realpath -m "$2")"; HDR="$3"; COVER="$(realpath "$4")"
DIR="$(cd "$(dirname "$0")" && pwd)"
TMP="$(mktemp -d)"
SOFFICE="${SOFFICE_PY:-/root/.claude/skills/synced/22a12918-27d4-4d21-b5e3-559b4837f450_9bad4d9a-ce49-4793-b76d-aa1d823109f0/docx/scripts/office/soffice.py}"
(cd /tmp && node "$DIR/md2docx.js" "$IN" "$TMP/pass1.docx" "$HDR" "$COVER")
(cd "$TMP" && python3 "$SOFFICE" --headless --convert-to pdf pass1.docx >/dev/null 2>&1)
python3 "$DIR/tocmap.py" "$TMP/pass1.pdf" "$TMP/pass1.docx.headings.json" "$TMP/toc.json"
(cd /tmp && node "$DIR/md2docx.js" "$IN" "$OUT" "$HDR" "$COVER" "$TMP/toc.json")
rm -f "$OUT.headings.json"
OUTDIR="$(dirname "$OUT")"
(cd "$OUTDIR" && python3 "$SOFFICE" --headless --convert-to pdf "$(basename "$OUT")" >/dev/null 2>&1)
echo "fertig: $OUT ($(pdfinfo "${OUT%.docx}.pdf" | awk '/Pages/{print $2}') Seiten)"

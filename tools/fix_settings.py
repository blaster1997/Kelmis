"""Bringt die Silbentrennungs-Einstellungen in word/settings.xml in die vom OOXML-Schema verlangte Reihenfolge."""
import re, shutil, sys, tempfile, zipfile

def fix(path):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                s = data.decode("utf-8")
                tags = {}
                for name in ("autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "evenAndOddHeaders"):
                    m = re.search(r"<w:%s\b[^>]*/>" % name, s)
                    if m:
                        tags[name] = m.group(0)
                        s = s.replace(m.group(0), "", 1)
                block = "".join(tags.get(n, "") for n in ("autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "evenAndOddHeaders"))
                s = s.replace("<w:compat>", block + "<w:compat>", 1)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)

if __name__ == "__main__":
    for p in sys.argv[1:]:
        fix(p)

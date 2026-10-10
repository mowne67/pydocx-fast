"""Run: python test_docx_speedups.py  -- patched output must match upstream byte-for-byte."""
import hashlib, io, subprocess, sys, time

SCRIPT = r'''
import sys, time, io, hashlib
if sys.argv[1] == "fast": import docx_speedups
from docx import Document
from docx.shared import Pt
N = int(sys.argv[2])
t0 = time.perf_counter()
d = Document()
d.add_heading("Title", 0)
for i in range(N):
    p = d.add_paragraph(f"Paragraph {i} ", style="List Bullet" if i % 7 == 0 else None)
    r = p.add_run("bold"); r.bold = True; r.font.size = Pt(11)
    p.add_run("\ttab\nbreak")
    if i % 1000 == 0:
        t = d.add_table(rows=3, cols=3)
        for c in t._cells: c.text = "cell"
        d.add_section()
        d.add_picture if False else None
d.sections[0].header.paragraphs[0].text = "header"
buf = io.BytesIO(); d.save(buf); t1 = time.perf_counter()
d2 = Document(io.BytesIO(buf.getvalue()))
text = "\n".join(p.text for p in d2.paragraphs)
cells = "|".join(c.text for t in d2.tables for c in t._cells)
t2 = time.perf_counter()
xml = d2.element.xml
print(hashlib.sha1((xml + text + cells).encode()).hexdigest(), f"{t1-t0:.2f}", f"{t2-t1:.2f}")
'''

def run(mode, n):
    out = subprocess.check_output([sys.executable, "-c", SCRIPT, mode, str(n)], text=True).split()
    return out[0], float(out[1]), float(out[2])

for n in (2_000, 50_000):
    slow, fast = run("slow", n), run("fast", n)
    assert slow[0] == fast[0], f"output differs at n={n}"
    print(f"n={n:>6}: write {slow[1]:6.2f}s -> {fast[1]:5.2f}s   read {slow[2]:5.2f}s -> {fast[2]:5.2f}s   (output identical)")

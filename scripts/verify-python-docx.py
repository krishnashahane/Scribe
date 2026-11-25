"""Independent check: open every exported DOCX with python-docx and summarise it."""
import sys, glob, docx
ok = True
for f in sorted(glob.glob(sys.argv[1] + '/*.docx')):
    try:
        d = docx.Document(f)
        paras = [p for p in d.paragraphs]
        imgs = len(d.inline_shapes)
        print(f"OK   {f.split('/')[-1]:32} paragraphs={len(paras):5} tables={len(d.tables):2} images={imgs} sections={len(d.sections)} title={d.core_properties.title!r}")
    except Exception as e:
        ok = False
        print(f"FAIL {f}: {e}")
sys.exit(0 if ok else 1)

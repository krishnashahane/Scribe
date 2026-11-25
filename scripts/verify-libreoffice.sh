#!/bin/sh
# Opens every exported DOCX in LibreOffice (headless), converts it to PDF and text,
# and compares page counts and text with the original fixture it was exported from.
set -u
SOFFICE=${SOFFICE:-/Applications/LibreOffice.app/Contents/MacOS/soffice}
OUT=tests/output/libreoffice
PROFILE="file://$(pwd)/tests/output/lo-profile"
mkdir -p "$OUT/exported" "$OUT/original"
convert() { "$SOFFICE" -env:UserInstallation="$PROFILE" --headless --convert-to "$1" --outdir "$2" "$3" >/dev/null 2>&1; }
pages() { [ -f "$1" ] && python3 -c "import re,sys;print(len(re.findall(rb'/Type\s*/Page[^s]',open(sys.argv[1],'rb').read())))" "$1" || echo "-"; }
fail=0
printf '%-34s %-8s %-11s %-11s %s\n' FILE OPENS 'PAGES(exp)' 'PAGES(orig)' TEXT
for f in tests/output/*.docx; do
  n=$(basename "$f" .docx)
  convert pdf "$OUT/exported" "$f"; convert txt:Text "$OUT/exported" "$f"
  ok=yes; [ -s "$OUT/exported/$n.pdf" ] || { ok=NO; fail=1; }
  orig=""; case "$n" in roundtrip-*) orig="tests/fixtures/${n#roundtrip-}.docx";; edited-simple) orig="tests/fixtures/simple.docx";; esac
  op="-"; text="(no original)"
  if [ -n "$orig" ] && [ -f "$orig" ]; then
    o=$(basename "$orig" .docx)
    convert pdf "$OUT/original" "$orig"; convert txt:Text "$OUT/original" "$orig"
    op=$(pages "$OUT/original/$o.pdf")
    if [ "${n#roundtrip-}" != "$n" ]; then
      if diff -q "$OUT/original/$o.txt" "$OUT/exported/$n.txt" >/dev/null 2>&1; then text="identical"; else text="DIFFERS"; fail=1; fi
    else text="(edited on purpose)"; fi
  fi
  printf '%-34s %-8s %-11s %-11s %s\n' "$n" "$ok" "$(pages "$OUT/exported/$n.pdf")" "$op" "$text"
done
exit $fail

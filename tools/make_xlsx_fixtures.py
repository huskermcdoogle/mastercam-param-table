"""Test fixtures for the xlsx reader (tests/xlsx_read_test.cpp).

Run:  python tools/make_xlsx_fixtures.py   (writes tests/fixtures/)

Everything here is made by Python's own zlib and zipfile, independent of the
C++ under test:

  deflate_<n>.bin / deflate_<n>.txt   raw DEFLATE at several levels, and the bytes
                                      it must inflate to
  excel_like.xlsx                     a workbook laid out the way EXCEL saves one:
                                      every part compressed, strings in
                                      sharedStrings.xml (one rich-text, one with
                                      Excel's _x000D_ escape for a carriage
                                      return, one with entities), a formula with its
                                      cached value, a formula returning text, a
                                      boolean, an error, a gap row, a group row
                                      above the column names.
"""
import os
import random
import zipfile
import zlib

here = os.path.dirname(os.path.abspath(__file__))
out = os.path.join(here, "..", "tests", "fixtures")
os.makedirs(out, exist_ok=True)


def raw_deflate(data, level):
    c = zlib.compressobj(level, zlib.DEFLATED, -15)
    return c.compress(data) + c.flush()


random.seed(7)
words = [b"feed", b"speed", b"0.008", b"per rev", b"coolant", b"10BAR", b"stock_x",
         b"<c r=\"A1\" t=\"s\"><v>12</v></c>", b"\r\n", b"op_idn"]
samples = {
    "1": b"",                                              # empty
    "2": b"a",                                             # one byte
    "3": b"".join(random.choice(words) for _ in range(5000)),  # repetitive, long matches
    "4": bytes(random.getrandbits(8) for _ in range(70000)),   # incompressible -> stored blocks
    "5": b"abc" * 30000,                                   # long overlapping copies
}
levels = {"1": 6, "2": 9, "3": 9, "4": 6, "5": 1}
for k, data in samples.items():
    open(os.path.join(out, f"deflate_{k}.bin"), "wb").write(raw_deflate(data, levels[k]))
    open(os.path.join(out, f"deflate_{k}.txt"), "wb").write(data)
# fixed-Huffman block: level 1 on a short input usually picks fixed codes
open(os.path.join(out, "deflate_6.bin"), "wb").write(raw_deflate(b"fixed huffman please", 1))
open(os.path.join(out, "deflate_6.txt"), "wb").write(b"fixed huffman please")

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
RNS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
shared = [
    "Identity",                                        # 0
    "op_idn",                                          # 1
    "type",                                            # 2
    "comment",                                         # 3
    "feed",                                            # 4
    "feed_mode",                                       # 5
    "manual_text",                                     # 6
    "ROUGH",                                           # 7
    "per rev",                                         # 8
    None,                                              # 9  rich text: "Datum " + "-A-"
    "M00_x000D_\nM01",                                 # 10 Excel's CR escape, then LF
    "R &amp; D &lt;1&gt;",                             # 11 entities
    "FINISH",                                          # 12
]
si = []
for s in shared:
    if s is None:
        si.append('<si><r><rPr><b/></rPr><t xml:space="preserve">Datum </t></r>'
                  '<r><t>-A-</t></r><rPh sb="0" eb="1"><t>IGNORED</t></rPh></si>')
    else:
        si.append(f'<si><t xml:space="preserve">{s}</t></si>')
sst = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<sst {NS} count="{len(si)}" uniqueCount="{len(si)}">{"".join(si)}</sst>')

sheet = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet {NS} {RNS}><dimension ref="A1:F6"/><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c></row>
<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="s"><v>2</v></c><c r="C2" t="s"><v>3</v></c><c r="D2" t="s"><v>4</v></c><c r="E2" t="s"><v>5</v></c><c r="F2" t="s"><v>6</v></c></row>
<row r="3"><c r="A3"><v>1</v></c><c r="B3" t="s"><v>7</v></c><c r="C3" t="s"><v>9</v></c><c r="D3"><v>1.0000000000000001E-2</v></c><c r="E3" t="s"><v>8</v></c><c r="F3" t="s"><v>10</v></c></row>
<row r="4" spans="1:6"/>
<row r="5"><c r="A5"><f>A3+6</f><v>7</v></c><c r="B5" t="s"><v>12</v></c><c r="C5" t="s"><v>11</v></c><c r="D5" t="str"><f>"0."&amp;"5"</f><v>0.5</v></c><c r="E5" t="b"><v>1</v></c><c r="F5" t="e"><v>#N/A</v></c></row>
<row r="6"><c r="A6" s="3"/><c r="B6" s="3"/></row>
</sheetData></worksheet>'''

workbook = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook {NS} {RNS}>'
            '<sheets><sheet name="Lathe params" sheetId="1" r:id="rId3"/></sheets></workbook>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="x/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="x/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="x/sharedStrings" Target="sharedStrings.xml"/>'
        '</Relationships>')

with zipfile.ZipFile(os.path.join(out, "excel_like.xlsx"), "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", "<Types/>")
    z.writestr("xl/workbook.xml", workbook)
    z.writestr("xl/_rels/workbook.xml.rels", rels)
    z.writestr("xl/sharedStrings.xml", sst)
    z.writestr("xl/worksheets/sheet1.xml", sheet)
print("fixtures written to", os.path.normpath(out))

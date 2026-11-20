"""Writes the .docx reproducer of storytold/craft-libs#2448 (VML half), from our own inputs and
without any word processor. Text is invented.

    python make-cell-shapes.py [out-dir]

cell-shapes.docx: A4, left margin 90 pt, compatibility mode 11. One paragraph, then three one-cell
bordered tables, like the boxes of a leaflet. A VML text box ("ICON n", 48 x 40 pt) is anchored in
the first paragraph of each cell, 60 pt left of the column and 4 pt below the paragraph's top,
wrapping top and bottom. Boxes 1 and 2 have o:allowincell="f" (Layout in table cell off); box 3 has
no o:allowincell (on, the default).

There is no .doc here: writing a .doc needs a compound-file writer. The .doc half
(fLayoutInCell in 0x03BF) is covered by the test in storytold/craft-libs#2450, which builds the
shape properties from the spec.
"""
import sys
import zipfile
from pathlib import Path

HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word"')
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CONTENT_TYPES = HEAD + """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="{REL}/settings" Target="settings.xml"/>
</Relationships>"""

STYLES = HEAD + f"""<w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""

SETTINGS = HEAD + (f'<w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" '
                   'w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat></w:settings>')

BODY_TEXT = ("Invented text for a box of a leaflet. It runs over a few lines, so that the cell has some "
             "height of its own and the place where its text starts is easy to see.")


def text_box(n, in_cell):
    allow = '' if in_cell else ' o:allowincell="f"'
    style = ("position:absolute;margin-left:-60pt;margin-top:4pt;width:48pt;height:40pt;z-index:1;"
             "mso-position-horizontal-relative:text;mso-position-vertical-relative:text")
    return (f'<w:r><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'
            '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
            f'<v:shape id="Box{n}" o:spid="_x0000_s10{n}" type="#_x0000_t202" style="{style}"{allow}>'
            f'<v:textbox><w:txbxContent><w:p><w:r><w:t>ICON {n}</w:t></w:r></w:p></w:txbxContent></v:textbox>'
            '<w10:wrap type="topAndBottom"/></v:shape></w:pict></w:r>')


def table(n, in_cell):
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for s in ("top", "left", "bottom", "right"))
    return (f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>'
            f'<w:p>{text_box(n, in_cell)}<w:r><w:t>BOX {n} TITLE</w:t></w:r></w:p>'
            f'<w:p><w:r><w:t>{BODY_TEXT}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>')


def document():
    body = ('<w:p><w:r><w:t>Synthetic leaflet: each box is a one-cell table; the label on its left is a text box '
            'anchored in the cell.</w:t></w:r></w:p><w:p/>')
    body += "".join(table(n, in_cell) for n, in_cell in ((1, False), (2, False), (3, True)))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return HEAD + f'<w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out / "cell-shapes.docx", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", document())

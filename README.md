# A positioned table in a .doc (comment on storytold/craft-libs#1823)

`make_positioned_table.py` writes the three `.doc` files here byte by byte, from [MS-DOC] and
[MS-CFB]; no application wrote them. The text is invented. Script and files: MIT OR Apache-2.0.

    python make_positioned_table.py

Each file has three paragraphs, a three-row table centred between the margins, five empty
paragraphs, and a paragraph with text whose first word is wider than the space left beside the
table. They differ only in where the table's position is stored:

| File | Position sprms |
|---|---|
| `positioned-table-row.doc` | in each row's TAPX: `sprmTPc`, `sprmTDxaAbs`, `sprmTDyaAbs`, `sprmTDxaFromText`, `sprmTDxaFromTextRight` |
| `positioned-table-cell-frames.doc` | only on the cell paragraphs: `sprmPPc`, `sprmPDxaAbs`, `sprmPDyaAbs`, `sprmPWr`, `sprmPDxaFromText` |
| `positioned-table-inline.doc` | none: an ordinary table |

## What Word shows

- **Row sprms:** the table floats. It sits a little below where the empty paragraphs start, and
  the five empty paragraphs sit beside it, at the height of its first rows, instead of below it.
  The paragraph with text starts right under the table.
- **Cell frame sprms only:** the table moves down as in the first file, but it does not float.
  The empty paragraphs and the paragraph with text come below it, as after an ordinary table.
- **Ordinary table:** the empty paragraphs and the paragraph with text follow the table.

So the row's own sprms are what make the table float.

## What craft-doc-binary does

The first file reads like the third: `RowInfo::decode` in `crates/doc-binary/src/table.rs` skips
the position sprms and `assemble` leaves `TableProps::float` unset. The table stays in the text
flow and the empty paragraphs push everything after them down.

[MS-DOC]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ccd7b486-7881-484c-a137-51170af7cc22
[MS-CFB]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b

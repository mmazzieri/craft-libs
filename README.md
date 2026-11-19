# Layout in table cell off: synthetic reproducer

Made with Word (Microsoft 365, Windows) by `make-cell-shapes.ps1`, through COM. The text is made up.

The document has three one-cell bordered tables, like the boxes of a one-page leaflet. A small text box ("ICON n") is anchored in the first paragraph of each cell. It is placed 60 pt left of the column, and wraps top and bottom. The document is in Word 2003 compatibility mode.

- Boxes 1 and 2 have **Layout in table cell** off.
- Box 3 has it on.

## Files

- `cell-shapes.doc`: Word 97–2003. Each text box has the Group Shape Boolean Properties (0x03BF):
  - boxes 1 and 2: `0x82000200`, so `fUsefLayoutInCell` (bit 31) is set and `fLayoutInCell` (bit 15) is clear;
  - box 3: `0x82008200`, so both bits are set.
- `cell-shapes.docx`: the same document saved by Word as `.docx`. The text boxes are VML; boxes 1 and 2 have `o:allowincell="f"` and box 3 has no `o:allowincell`.
- `cell-shapes-word.pdf`: Word's PDF.
  - Boxes 1 and 2 sit left of their table.
  - Box 3 is moved inside its cell, and the cell's text starts below it.
- `export-0.18.1.pdf`: WordCraft 0.18.1, `wordcraft-cli convert cell-shapes.docx export-0.18.1.pdf`.
- `import-0.18.1.pdf`: WordCraft 0.18.1, `.doc` → `.docx` → `.pdf`.
  - In both, all three text boxes are inside their cells, and every cell's text starts below them.
- `export-0.15.0.pdf`: WordCraft 0.15.0, before the layout kept a cell's floats in the cell (wordcraft#4545). All three text boxes are outside, box 3 too.
- `side-by-side.png`: page tops of Word, 0.15.0, 0.18.1 export and 0.18.1 import, left to right.
- `export-with-fix.pdf`, `import-with-fix.pdf`: wordcraft `main` (6ef13b2c) built against craft-libs 93dac78c with the two reader fixes on top. Boxes 1 and 2 are outside their tables and box 3 is inside, as in Word.
- `side-by-side-with-fix.png`: page tops of Word, wordcraft `main` export, and the fixed export and import, left to right.
- `probe/`: the same `.doc` with the bits changed by hand:
  - box 1 `0x02000200`, both bits clear;
  - box 3 `0x02008200`, only `fLayoutInCell` set.

  Word's `.docx` and PDF of it are included. Word draws boxes 1 and 3 in their cells and writes no `o:allowincell` for them. So without the use bit the setting is on, its default.

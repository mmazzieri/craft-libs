# Synthetic reproducer: a `.doc` table cell's text direction

Everything here is invented and was made by `make-turned-cells-doc.ps1` with Word 16 (COM) on Windows.

`turned-cells.doc` and `turned-cells.docx` are the same table, saved by Word in both formats. It has two rows of three cells, each row exactly 90 pt high:

| | cell 1 | cell 2 | cell 3 |
|---|---|---|---|
| row 1 | horizontal | upward (`btLr`) | downward (`tbRl`) |
| row 2 | downward (`tbRl`) | horizontal | upward (`btLr`) |

**What Word writes in the `.doc`:**

- Bits 2–4 (`textFlow`) of each cell's TC80 `TCGRF` in the row's `sprmTDefTable`: 3 on the `btLr` cells, 1 on the `tbRl` cells, 0 on the others.
- One `sprmTTextFlow` (0x7629) per turned cell, with a 4-byte operand: `itcFirst`, `itcLim`, then the same value as a u16. Row 1 has `01 02 03 00` and `02 03 01 00`.

**The `.docx`** has `w:textDirection` `btLr` / `tbRl` on exactly those four cells.

**Files:**

- `turned-cells-word.pdf`: Word's export.
- `turned-cells-base.pdf`: `wordcraft-cli convert turned-cells.doc`, built from wordcraft `main` f8ec3cc0 with craft-libs at its pin 93dac78c plus #2414. Every cell comes out horizontal, and the `.docx` it writes has no `w:textDirection`.
- `turned-cells-fix.pdf`: the same build plus the fix. The four cells are turned, and the `.docx` it writes has the same four `w:textDirection` as Word's.
- `side-by-side.png`: Word, base and fix, left to right.

The fix renders the columns narrower than Word does. That comes from the cell widths of a `.doc`, which #1856 and #2117 cover, and has nothing to do with the text direction.

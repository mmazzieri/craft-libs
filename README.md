# Synthetic reproducer: a `0x81` bold toggle in a bold paragraph style (.doc)

`make-bold-toggle.ps1` builds `bold-toggle.doc` with Word 2016 via COM. The text is invented.

- A paragraph style `Bold Body`, based on Normal, is bold. Its CHPX is `35 08 81`.
- One paragraph in the body and one inside a one-cell table use it.
- In each, the middle stretch is set not bold. Word stores it as `sprmCFBold` `0x81` (`35 08 81`, "the opposite of the style").

| file | what it is |
|---|---|
| `bold-toggle.doc` | the file Word saved |
| `bold-toggle.docx`, `bold-toggle.pdf` | Word's own save of the same document: the middle stretches are not bold (`<w:b w:val="0"/>`) |
| `read-craft-libs-main.docx` | `craft_doc_binary::read` then `craft_docx::write`, craft-libs `main` 8048f8ef: both middle stretches come out **bold** (`<w:b/>`) |
| `read-pr-1804.docx` | the same with the head of #1804 (05124b47): not bold, in the body and in the cell, as Word |

No personal properties: checked on the DOC, the DOCX and the PDF.

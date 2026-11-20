# A table cell's text direction in a .doc (storytold/craft-libs#2529)

This branch held a `.doc` and a `.docx` saved by Word through a COM script, Word's PDF, WordCraft's
PDFs and a picture comparing them. They are removed: under the clean-room standard a fixture is
written by a committed script from our own inputs, and an agent does not drive Word to make one.

A `.doc` cannot be written here without a compound-file writer, so there is no replacement file.
The reader change is covered by the tests of storytold/craft-libs#2541, which build the TC80s and
the `sprmTTextFlow` operands from [MS-DOC]:

- `text_direction_kept_in_tc80s_is_read`
- `an_explicit_text_flow_overrides_the_tc80`

The layout of the records, from [MS-DOC]: `textFlow` is bits 2–4 of the `TCGRF`; `sprmTTextFlow`
(0x7629) takes `itcFirst`, `itcLim` and the value as a u16.

[MS-DOC]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/3dd0761d-34df-4b7a-8d62-d74e23b2d1e0

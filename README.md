# Bold toggled against the paragraph style in a .doc (comment on storytold/craft-libs#1804)

This branch held a `.doc` and a `.docx` saved by Word through a COM script, Word's PDF and the
output of two builds. They are removed: under the clean-room standard a fixture is written by a
committed script from our own inputs, and an agent does not drive Word to make one.

A `.doc` cannot be written here without a compound-file writer, so there is no replacement file.
The case it showed, in terms of [MS-DOC]: a paragraph style whose CHPX turns bold on
(`sprmCFBold`, `35 08 01`), and a run in a paragraph of that style, in the body or in a table
cell, whose CHPX holds `sprmCFBold` with the operand 0x81 (`35 08 81`, "the opposite of the
style"). The run is not bold. Resolving 0x81 before the paragraph's style is known reads it as
bold.

[MS-DOC]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/3dd0761d-34df-4b7a-8d62-d74e23b2d1e0

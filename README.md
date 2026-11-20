# Vertical merge and alignment kept only in TC80s (storytold/craft-libs#2413)

This branch held a `.doc` and a `.docx` saved by Word through a COM script, Word's PDF and a
rendering. They are removed: under the clean-room standard a fixture is written by a committed
script from our own inputs, and an agent does not drive Word to make one.

A `.doc` cannot be written here without a compound-file writer, so there is no replacement file.
The reader change is covered by the tests of storytold/craft-libs#2414, which build the
`sprmTDefTable` operands and their TC80s from [MS-DOC]:

- `vertical_merge_and_alignment_kept_only_in_tc80s_are_read`
- `an_explicit_vertical_merge_or_alignment_overrides_the_tc80`
- `tc80_flags_are_read_only_from_whole_tc80s`

The layout of the record, from [MS-DOC] (`TC80`, `TCGRF`): `vertMerge` is bits 5–6 (3 starts a
merge, 1 continues it) and `vertAlign` is bits 7–8 (1 centre, 2 bottom).

[MS-DOC]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/3dd0761d-34df-4b7a-8d62-d74e23b2d1e0

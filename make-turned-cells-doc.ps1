# Synthetic reproducer: a one-row table whose three cells run their text horizontally, upward
# (btLr) and downward (tbRl), plus a second row with the directions swapped. Saved by Word as
# .doc and .docx. Invented text.
param([string] $Out = $PSScriptRoot)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $d = $word.Documents.Add()
    $d.RemovePersonalInformation = $true
    $d.Content.Text = 'Before the table.'
    $d.Paragraphs(1).Range.InsertParagraphAfter()
    $t = $d.Tables.Add($d.Paragraphs(2).Range, 2, 3)
    $t.Borders.Enable = $true
    $t.Rows.HeightRule = 2   # wdRowHeightExactly
    $t.Rows.Height = 90
    # wdTextOrientationHorizontal 0, Upward 2, Downward 3
    $dirs = @(@(0, 2, 3), @(3, 0, 2))
    for ($r = 1; $r -le 2; $r++) {
        for ($c = 1; $c -le 3; $c++) {
            $cell = $t.Cell($r, $c)
            $cell.Range.Text = "Invented cell $r.$c"
            $cell.Range.Orientation = $dirs[$r - 1][$c - 1]
        }
    }
    $d.SaveAs2((Join-Path $Out 'turned-cells.docx'), 16)
    $d.SaveAs2((Join-Path $Out 'turned-cells.doc'), 0)
    $d.ExportAsFixedFormat((Join-Path $Out 'turned-cells-word.pdf'), 17)
    $d.Close(0)
} finally { $word.Quit() }

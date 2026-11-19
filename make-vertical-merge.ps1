# Builds vertical-merge.doc and vertical-merge.docx with Word (Windows, COM).
# A 6 x 3 table: column 1 merged over rows 1-4 and 5-6, column 2 over rows 1-2, 3-4 and 5-6;
# column 3 unmerged. Cells (1,1), (1,2) and (1,3) are centred vertically, (2,3) bottom-aligned.
# All text is invented.
param([string]$Out = $PSScriptRoot)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $d = $word.Documents.Add()
    $d.RemovePersonalInformation = $true
    $d.Range().Text = 'Synthetic header table with vertically merged cells'
    $d.Content.InsertParagraphAfter()
    $t = $d.Tables.Add($d.Paragraphs($d.Paragraphs.Count).Range, 6, 3)
    $t.Borders.Enable = $true
    $t.Columns(1).Width = 120; $t.Columns(2).Width = 220; $t.Columns(3).Width = 120
    for ($i = 1; $i -le 6; $i++) { $t.Cell($i, 3).Range.Text = "Item $i" }
    $t.Cell(1, 1).Range.Text = 'Logo'
    $t.Cell(5, 1).Range.Text = 'Department'
    $t.Cell(1, 2).Range.Text = 'Title'
    $t.Cell(3, 2).Range.Text = 'Subtitle'
    $t.Cell(1, 3).VerticalAlignment = 1 # centre
    $t.Cell(2, 3).VerticalAlignment = 3 # bottom
    # Bottom-up, so the cell indexes above stay valid.
    $t.Cell(5, 2).Merge($t.Cell(6, 2)); $t.Cell(5, 1).Merge($t.Cell(6, 1))
    $t.Cell(3, 2).Merge($t.Cell(4, 2)); $t.Cell(1, 2).Merge($t.Cell(2, 2))
    $t.Cell(1, 1).Merge($t.Cell(4, 1))
    $t.Cell(1, 1).VerticalAlignment = 1
    $t.Cell(1, 2).VerticalAlignment = 1
    $d.Content.InsertParagraphAfter()
    $d.Content.InsertAfter('Text after the table.')
    $d.SaveAs2((Join-Path $Out 'vertical-merge.doc'), 0)   # wdFormatDocument97
    $d.SaveAs2((Join-Path $Out 'vertical-merge.docx'), 16) # wdFormatXMLDocument
    $d.Close($false)
} finally {
    $word.Quit()
}

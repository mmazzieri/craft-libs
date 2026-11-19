# Builds a synthetic reproducer with Word (COM): three one-cell bordered tables, each with a
# small text box anchored in the cell's first paragraph, offset to the left of the table
# (outside the cell), wrapping top and bottom. "Layout in table cell" is off in tables 1-2
# and on in table 3. Saved in Word 97-2003 format (.doc), as .docx in Word 2003
# compatibility mode (VML shapes, o:allowincell), and as PDF (Word's own rendering).
param([string]$Out = $PSScriptRoot)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Add()
    $doc.PageSetup.LeftMargin = 90   # room on the left for the boxes
    $sel = $word.Selection
    $sel.TypeText('Synthetic leaflet: each box is a one-cell table; the label on its left is a text box anchored in the cell.')
    $sel.TypeParagraph()
    $lic = @($false, $false, $true)
    for ($i = 0; $i -lt 3; $i++) {
        $sel.TypeParagraph()
        $range = $sel.Range
        $t = $doc.Tables.Add($range, 1, 1)
        $t.Borders.Enable = $true
        $cell = $t.Cell(1, 1).Range
        $cell.Text = "BOX $($i + 1) TITLE`rLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris."
        $anchor = $t.Cell(1, 1).Range.Paragraphs(1).Range
        $tb = $doc.Shapes.AddTextbox(1, -60, 4, 48, 40, $anchor)
        $tb.TextFrame.TextRange.Text = "ICON $($i + 1)"
        $tb.LayoutInCell = $lic[$i]          # first: Word keeps an in-cell box inside its cell
        $tb.WrapFormat.Type = 4              # wdWrapTopBottom
        $tb.RelativeHorizontalPosition = 2   # wdRelativeHorizontalPositionColumn
        $tb.RelativeVerticalPosition = 2     # wdRelativeVerticalPositionParagraph
        $tb.Left = -60
        $tb.Top = 4
        $sel.EndKey(6) | Out-Null            # wdStory
    }
    $doc.SetCompatibilityMode(11)
    $doc.RemovePersonalInformation = $true
    $doc.SaveAs2((Join-Path $Out 'cell-shapes.doc'), 0)          # wdFormatDocument97
    $doc.SaveAs2((Join-Path $Out 'cell-shapes.docx'), 16, $false, '', $true, '', $false, $false, $false, $false, $false, 0, $false, $false, 0, $false, 11)
    $doc.ExportAsFixedFormat((Join-Path $Out 'cell-shapes-word.pdf'), 17)
    $doc.Close($false)
} finally { $word.Quit() }

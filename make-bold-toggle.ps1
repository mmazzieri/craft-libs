# Builds bold-toggle.doc with Word via COM: a bold paragraph style, one paragraph in the body
# and the same paragraph inside a one-cell table, with a middle stretch set not bold.
param([string]$Out = (Join-Path $PSScriptRoot 'bold-toggle'))
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Add()
    $st = $doc.Styles.Add('Bold Body', 1)
    $normal = $doc.Styles.Item(-1).NameLocal
    $st.BaseStyle = $normal
    $st.Font.Bold = $true

    $sel = $word.Selection
    $sel.Style = $doc.Styles.Item('Bold Body')
    $sel.TypeText('BODY bold start, ')
    $sel.Font.Bold = $false
    $sel.TypeText('body plain middle')
    $sel.Font.Bold = $true
    $sel.TypeText(', body bold end.')
    $sel.TypeParagraph()
    $sel.Style = $doc.Styles.Item(-1)
    $sel.TypeParagraph()

    $tbl = $doc.Tables.Add($sel.Range, 1, 1)
    $tbl.Borders.Enable = $true
    $cell = $tbl.Cell(1, 1).Range
    $cell.Style = $doc.Styles.Item('Bold Body')
    $tbl.Cell(1, 1).Select()
    $sel.Collapse(1)
    $sel.TypeText('CELL bold start, ')
    $sel.Font.Bold = $false
    $sel.TypeText('cell plain middle')
    $sel.Font.Bold = $true
    $sel.TypeText(', cell bold end.')

    # No personal properties in the file.
    $doc.RemoveDocumentInformation(99)
    $doc.SaveAs2("$Out.doc", 0)
    $doc.SaveAs2("$Out.docx", 16)
    $doc.SaveAs2("$Out.pdf", 17)
    $doc.Close($false)
} finally {
    $word.Quit()
}

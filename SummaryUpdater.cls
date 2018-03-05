VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummaryUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ISummaryUpdater

Private Function ISummaryUpdater_ClearSheet(ws As Worksheet) As Variant

End Function

Private Function ISummaryUpdater_FitAllColumns(ws As Worksheet) As Variant
    ws.Cells.Select
    ws.Cells.EntireColumn.AutoFit
    ws.Cells(1, 1).Select
End Function

Private Function ISummaryUpdater_GetVersion() As String
    ISummaryUpdater_GetVersion = "1.0"
End Function

Private Sub ISummaryUpdater_Update(engine As RuleEngine, sheet As Worksheet)
    Dim row As RowInfo
    Dim pRule As rule
    Dim offset As Integer
    
    Dim ColHeaders As New Collection
    ColHeaders.Add ("PriceL")
    ColHeaders.Add ("PriceB")
    
    Dim colHeader1 As String
    Dim colHeader2 As String
    Dim colHeader3 As String
    Dim colHeader4 As String
    
    colHeader1 = sheet.Range("B1").Value
    colHeader2 = sheet.Range("C1").Value
    ' go through all available rules
    For Each pRule In engine.rules
        sheet.Range("A2").offset(offset, 0) = pRule.Category
        
        ' Get total for each rule
        Dim pTotal1 As Double: pTotal1 = 0
        Dim pTotal2 As Double: pTotal2 = 0
        For Each row In engine.Rows
            If row.Category = pRule.Category Then
                pTotal1 = pTotal1 + row.Columns(colHeader1)
                pTotal2 = pTotal2 + row.Columns(colHeader2)
            End If
        Next
        sheet.Range("B2").offset(offset, 0) = pTotal1
        sheet.Range("B2").offset(offset, 0).Interior.ColorIndex = 0
        sheet.Range("C2").offset(offset, 0) = pTotal2
        sheet.Range("C2").offset(offset, 0).Interior.ColorIndex = 0
        
        offset = offset + 1
    Next
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummaryUpdaterAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISummaryUpdater

Private Function ISummaryUpdater_ClearSheet(ws As Worksheet) As Variant
ws.Range("A2:ZZ6000").ClearContents
End Function

Private Function ISummaryUpdater_FitAllColumns(ws As Worksheet) As Variant
    ws.Activate
    ws.Cells.Select
    ws.Cells.EntireColumn.AutoFit
    ws.Cells(1, 1).Select
End Function

Private Function ISummaryUpdater_GetVersion() As String

End Function

Private Sub ISummaryUpdater_Update(engine As RuleEngine, sheet As Worksheet)
    Dim row As RowInfo
    Dim pRule As rule
    Dim offset As Integer
    
    
    Dim r As Range
    Dim c As Range
    Set r = sheet.Range(sheet.Range("B1"), sheet.Range("B1").End(xlToRight))
    'r.Select
    
 
    offset = 1
    ' go through all available rules
    For Each pRule In engine.rules
        sheet.Range("A2").offset(offset - 1, 0) = pRule.Category

        For Each row In engine.Rows
            If row.Category = pRule.Category Then
               
                For Each c In r
                    c.offset(offset, 0).Value = c.offset(offset, 0).Value + row.Columns(c.Value)
                Next
            End If
        Next

        
        offset = offset + 1
    Next
        
End Sub

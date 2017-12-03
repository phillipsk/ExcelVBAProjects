Attribute VB_Name = "Main"
Option Explicit

Sub Execute()

    Dim row As RowInfo
    'Dim rule As rule
            
    Dim engine As New RuleEngine
    Set engine.Rows = ReadAllRows()
    Set engine.Rules = ReadAllRules()
    Call engine.Apply
    
'    For Each row In engine.Rows
'        Debug.Print row.Columns("Id") & "=" & row.Category
'    Next
    

End Sub



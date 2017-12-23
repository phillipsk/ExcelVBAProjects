Attribute VB_Name = "Main"
Option Explicit

Sub Main()

    
    ' Read values from worksheet and rules + apply rules
    
    Dim row As RowInfo
    Dim engine As New RuleEngine
    Set engine.Rows = ReadAllRows()
    Set engine.Rules = ReadAllRules()
    Call engine.Apply
    

    
    
    ' Updating worksheet based on results
    
    Dim ws As Worksheet
    Set ws = Worksheets("Groceries")
    Dim r As Range
    Dim Total As Double
    Dim fruitTotal As Double
    fruitTotal = 0
    Total = 0
    
    Dim totalFormula As String
    totalFormula = "=SUM("
    ' C4,C5,C8
    For Each row In engine.Rows
        Set r = ws.Range("A" & row.RowNumber)
        
        ws.Range("L" & row.RowNumber).Value = row.Category
        
        'Debug.Print row.Columns("Id") & "=" & row.Category
        r.Interior.ColorIndex = 0
        
        
        
        If row.Category = "RuleIV" Then
            r.Interior.ColorIndex = 6
            Total = Total + row.Columns("PriceL")
            totalFormula = totalFormula & ws.Range("C" & row.RowNumber).Address & ","
        End If
        
        If row.Category = "FruitI" Then
            r.Interior.ColorIndex = 4
            fruitTotal = fruitTotal + row.Columns("PriceL")
        End If
        
    Next
    'MsgBox total
    'row.PriceL
    ws.Range("C19").Value = Total
    ws.Range("C20").Value = fruitTotal
    totalFormula = Left(totalFormula, Len(totalFormula) - 1)
    totalFormula = totalFormula & ")"
    
    ' check if there were no matching cells for total formula
    If totalFormula <> "=SUM)" Then
        ws.Range("C21").Value = totalFormula
    Else
        ws.Range("C21").Value = ""
    End If
End Sub



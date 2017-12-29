Attribute VB_Name = "Main"
Option Explicit

Sub Main()

'Tasks
'======
'
'1. Complete Rule Engine <-
'2. Find edge cases where it is not working.
'3. Make Data and Rules worksheets dynamic
'
'4. Grouping using brackets (a = b or b = c) and b = d (Operator Prescedence)
'
'5*. If (Rule1  OR Rule2) = Rule8

'TODO
'=====
'
'1. [x] Get existing rule results (added 2 new functions to RuleEngine class)
'2. [x] AND/OR operators
'3. [x] Make Data and Rules worksheets dynamic
'4. [x] Clear category column before applying rules
'5. [ ] Finish remaining operators like <>
'6. [ ] Pass Workbook name dynamically
'
'
'Additional ideas
'=================
'
'1. Add arithmetic operators like +/-
'2. Try Macro record to general basic macros

    
    ' Read values from worksheet and rules + apply rules
    
    Dim row As RowInfo
    Dim engine As New RuleEngine
    Set engine.Rows = ReadAllRows("Groceries")
    Set engine.rules = ReadAllRules("Rules")
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
        
        'COLUMN WHERE CATEGORY IS PRINTED
        
        ws.Range("M" & row.RowNumber).ClearContents
        ws.Range("M" & row.RowNumber).Value = row.Category
        
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



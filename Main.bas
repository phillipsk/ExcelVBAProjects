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

'Tasks
'=====
'
'1. [ ] Add NotEq <> operator
'2. [ ] Understand Str and CStr functions in VBA
'
'
'
'Others
'======
'1. [ ] Simplify Select Case in RuleEngine
'2. [ ] Improve AND/OR logic


    
    ' Read values from worksheet and rules + apply rules
        
    Dim engine As New RuleEngine
    Set engine.Rows = ReadAllRows("Groceries")
    Set engine.rules = ReadAllRules("Rules")
    Call engine.Apply
    
    ' Update category column in Groceries worksheet.
    UpdateCategory engine
    
    UpdateSummary engine
    
    ' MsgBox (GetTotal(New totalProvider, engine))

End Sub

' Updates the category column with update rule results.
Private Sub UpdateCategory(engine As RuleEngine)

    ' Updating worksheet based on results
    Dim row As RowInfo
    Dim ws As Worksheet
    Set ws = Worksheets("Groceries")
    Dim r As Range
    Dim Total As Double: Total = 0

    For Each row In engine.Rows
        Set r = ws.Range("A" & row.RowNumber)
        'COLUMN WHERE CATEGORY IS PRINTED
        ws.Range("M" & row.RowNumber).ClearContents
        ws.Range("M" & row.RowNumber).Value = row.Category
        
        'Debug.Print row.Columns("Id") & "=" & row.Category
        r.Interior.ColorIndex = 0

    Next
End Sub

Private Sub UpdateSummary(engine As RuleEngine)
    Dim row As RowInfo
    Dim pRule As rule
    Dim offset As Integer
    Dim colHeaders As New Collection
    colHeaders.Add ("PriceL")
    colHeaders.Add ("PriceB")
    
    Dim colHeader1 As String
    Dim colHeader2 As String
    Dim colHeader3 As String
    Dim colHeader4 As String
    
    colHeader1 = Worksheets("Summary").Range("B1").Value
    colHeader2 = Worksheets("Summary").Range("C1").Value
    ' go through all available rules
    For Each pRule In engine.rules
        Worksheets("Summary").Range("A2").offset(offset, 0) = pRule.Category
        
        ' Get total for each rule
        Dim pTotal1 As Double: pTotal1 = 0
        Dim pTotal2 As Double: pTotal2 = 0
        For Each row In engine.Rows
            If row.Category = pRule.Category Then
                pTotal1 = pTotal1 + row.Columns(colHeader1)
                pTotal2 = pTotal2 + row.Columns(colHeader2)
            End If
        Next
        Worksheets("Summary").Range("B2").offset(offset, 0) = pTotal1
        Worksheets("Summary").Range("C2").offset(offset, 0) = pTotal2
        
        offset = offset + 1
    Next
End Sub

'Private Function GetTotal(totalProvider As totalProvider, engine As RuleEngine) As Double
'    GetTotal = totalProvider.GetTotal(engine)
'End Function


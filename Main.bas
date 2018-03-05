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
    
    'UpdateSummary engine
    Dim updater As New summaryUpdater
    Dim updaterColorful As New SummaryUpdaterColorful
    Dim updaterAdd As New SummaryUpdaterAdd
    Dim updaterSubtract As New SummaryUpdaterSubtract
    
    
    UpdateSummary updaterAdd, engine, Worksheets("Summary")
    UpdateSummary updaterSubtract, engine, Worksheets("SummarySubtract")
    'Or pass a variable
    'UpdateSummary updaterColorful, engine, Worksheets("Summary")
    
    'MsgBox (GetTotal(New totalProvider, engine))

End Sub

' Updates the category column with update rule results.
Private Sub UpdateCategory(engine As RuleEngine)

    ' Updating worksheet based on results
    Dim row As RowInfo
    Dim ws As Worksheet
    Set ws = Worksheets("Groceries")
    Dim r As Range
    Dim Total As Double: Total = 0

    Dim colLetter As String
    
    colLetter = Col_Letter(ws.Range("A1").End(xlToRight).Column)
    For Each row In engine.Rows
        Set r = ws.Range("A" & row.RowNumber)
        'COLUMN WHERE CATEGORY IS PRINTED
        
        ws.Range(colLetter & row.RowNumber).ClearContents
        ws.Range(colLetter & row.RowNumber).Value = row.Category
        
        'Debug.Print row.Columns("Id") & "=" & row.Category
        r.Interior.ColorIndex = 0

    Next
End Sub

Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

'Private Function GetTotal(totalProvider As totalProvider, engine As RuleEngine) As Double
'    GetTotal = totalProvider.GetTotal(engine)
'End Function

'Sub Test()
'
'    Dim cat As New cat
'    Dim dog As New dog
'
'    AnimalSpeak cat
'    AnimalSpeak dog
'
'End Sub
'
'Sub AnimalSpeak(a As IAnimal)
'    a.Speak
'End Sub

Sub UpdateSummary(summaryUpdater As ISummaryUpdater, engine As RuleEngine, sheet As Worksheet)
    summaryUpdater.ClearSheet sheet
    summaryUpdater.Update engine, sheet
    summaryUpdater.FitAllColumns sheet
    'Debug.Print summaryUpdater.GetVersion()
End Sub



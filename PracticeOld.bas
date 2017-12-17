Attribute VB_Name = "PracticeOld"
Sub Practice()
'Sheets("Test file").Select
Dim varPrac As Integer

LastCellBeforeBlankInColumn
LastCellInColumn
LastCellBeforeBlankInRow
LastCellInRow
    Application.SendKeys "^g ^a {DEL}"
End Sub
'Find the last used cell, before a blank in a Column:
Sub LastCellBeforeBlankInColumn()

varPrac = Range("A2").End(xlDown).row
Debug.Print "LastCellBeforeBlankInColumn"; varPrac
End Sub
'Find the very last used cell in a Column:
Sub LastCellInColumn()

varPrac = Range("A65536").End(xlUp).row
Debug.Print "LastCellInColumn"; varPrac
End Sub
'Find the last cell, before a blank in a Row:
Sub LastCellBeforeBlankInRow()

varPrac = Range("A2").End(xlToRight).Column
Debug.Print "LastCellBeforeBlankInRow"; varPrac
End Sub
'Find the very last used cell in a Row:
Sub LastCellInRow()

varPrac = Range("IV2").End(xlToLeft).Column
Debug.Print "LastCellInRow"; varPrac
End Sub
Sub CompareStrings()
Dim A As String, B() As Variant
'Sheets("Practice").Select

A = "Cat"
B = Application.Transpose(Range("A1:A8").Value)

For i = 1 To 8

    MsgBox A = B(i)

    MsgBox A Like B(i)

    MsgBox StrComp(A, B(i)) = 0

    MsgBox "Cat" = B(i)

Next

End Sub


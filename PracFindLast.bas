Attribute VB_Name = "PracFindLast"
Sub MainPractice()
Dim pracInt As Integer
Debug.Print
Debug.Print Range("A2").End(xlDown).Column; " "; "Last BLANK cell in a Column:"
Debug.Print Range("A50").End(xlUp).Column; " "; "Last OCCUPIED cell in a Column:"
Debug.Print Range("A2").End(xlToRight).row; " "; "Last BLANK cell in a Row:"
Debug.Print Range("IV2").End(xlToLeft).row; " "; "Last OCCUPIED cell in a Row:"
Debug.Print
Debug.Print Range("A2").End(xlDown).Address(0, 0); " "; "Last BLANK cell in a Column:"
Debug.Print Range("A50").End(xlUp).Address(0, 0); " "; "Last OCCUPIED cell in a Column:"
Debug.Print Range("A2").End(xlToRight).Address(0, 0); " "; "Last BLANK cell in a Row:"
Debug.Print Range("IV2").End(xlToLeft).Address(0, 0); " "; "Last OCCUPIED cell in a Row:"
Application.SendKeys "^g ^a {DEL}"
End Sub

Attribute VB_Name = "Utility"

' example: converts "=" to Operators.Eq
Function GetOperatorValue(op As String) As Operators
    Select Case UCase(op)
        Case "="
            GetOperatorValue = Eq
        Case ">"
            GetOperatorValue = Gt
        Case "<"
            GetOperatorValue = Lt
        Case ">="
            GetOperatorValue = GtEq
        Case "<="
            GetOperatorValue = LtEq
        Case "<>"
            GetOperatorValue = NoEQ
        Case "+"
            GetOperatorValue = Add
        Case "-"
            GetOperatorValue = Subtract
        Case "++"
            GetOperatorValue = Pos
        Case "--"
            GetOperatorValue = Neg
        Case "AND"
            GetOperatorValue = [And]
        Case "OR"
            GetOperatorValue = [Or]
    End Select
End Function

Function IsOperator(txt As String) As Boolean
    Select Case txt
        Case "="
            IsOperator = True
        Case ">"
            IsOperator = True
        Case "<"
            IsOperator = True
        Case ">="
            IsOperator = True
        Case "<="
            IsOperator = True
        Case "<>"
            IsOperator = True
        Case "+"
            IsOperator = True
        Case "-"
            IsOperator = True
        Case "++"
            IsOperator = True
        Case "--"
            IsOperator = True
        Case Else
            IsOperator = False
    End Select
End Function

Function IsLogicalOperator(txt As String) As Boolean
    Select Case txt
        Case "AND"
            IsLogicalOperator = True
        Case "OR"
            IsLogicalOperator = True
        Else
            IsKeyword = False
    End Select

End Function

Function IsKeyword(txt As String) As Boolean
    Select Case txt
        Case "RETURN"
            IsKeyword = True
        Else
            IsKeyword = False
    End Select
End Function

Function IsRowInfo(obj As Variant) As Boolean
    IsRowInfo = (TypeName(obj) = "RowInfo") ' true / false

' Output in Immediate Window.
'Set ri = New RowInfo
'Set ri2 = New RowInfo
'Debug.Print IsRowInfo(ri2)
'True
End Function

Function IsFoundInDictionary(d As Dictionary, k As Variant) As Boolean
    IsFoundInDictionary = d.Exists(k) ' Checks if key exists in dictionary.
End Function

Function IsRomNum(c As String) As Boolean
'    I   V   X   L   C   D   M
    Select Case c
        Case "I"
            IsRomNum = True
        Case "II"
            IsRomNum = True
        Case "III"
            IsRomNum = True
        Case "IV"
            IsRomNum = True
        Case "V"
            IsRomNum = True
        Case "VI"
            IsRomNum = True
        Case "VII"
            IsRomNum = True
        Case "VIII"
            IsRomNum = True
        Case "IV"
            IsRomNum = True
        Case "V"
            IsRomNum = True
        Case Else
            IsRomNum = False
    End Select

End Function
' Return the Arabic version of this number.
Function RomanToArabic(ByVal roman As String) As _
    Long
Dim i As Integer
Dim ch As String
Dim result As Long
Dim new_value As Long
Dim old_value As Long

    roman = UCase$(roman)
    old_value = 1000

    For i = 1 To Len(roman)
        ' See what the next character is worth.
        ch = Mid$(roman, i, 1)
        Select Case ch
            Case "I"
                new_value = 1
            Case "V"
                new_value = 5
            Case "X"
                new_value = 10
            Case "L"
                new_value = 50
            Case "C"
                new_value = 100
            Case "D"
                new_value = 500
            Case "M"
                new_value = 1000
        End Select

        ' See if this character is bigger
        ' than the previous one.
        If new_value > old_value Then
            ' The new value > the previous one.
            ' Add this value to the result
            ' and subtract the previous one twice.
            result = result + new_value - 2 * old_value
        Else
            ' The new value <= the previous one.
            ' Add it to the result.
            result = result + new_value
        End If

        old_value = new_value
    Next i

    RomanToArabic = result
End Function
' Formats a number as a roman numeral.
' Author: Christian d'Heureuse (www.source-code.biz)
Public Function FormatRoman(ByVal n As Integer) As String
   If n = 0 Then FormatRoman = "0": Exit Function
      ' There is no roman symbol for 0, but we don't want to return an empty string.
   Const r = "IVXLCDM"              ' roman symbols
   Dim i As Integer: i = Abs(n)
   Dim s As String, p As Integer
   For p = 1 To 5 Step 2
      Dim d As Integer: d = i Mod 10: i = i \ 10
      Select Case d                 ' format a decimal digit
         Case 0 To 3: s = String(d, Mid(r, p, 1)) & s
         Case 4:      s = Mid(r, p, 2) & s
         Case 5 To 8: s = Mid(r, p + 1, 1) & String(d - 5, Mid(r, p, 1)) & s
         Case 9:      s = Mid(r, p, 1) & Mid(r, p + 2, 1) & s
         End Select
      Next
   s = String(i, "M") & s           ' format thousands
   If n < 0 Then s = "-" & s        ' insert sign if negative (non-standard)
   FormatRoman = s
   End Function


Function IsRule(columnName As String) As Boolean
    IsRule = UCase(Left(columnName, 4)) = "RULE"
End Function



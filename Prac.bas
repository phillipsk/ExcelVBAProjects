Attribute VB_Name = "Prac"
Sub Prac()
CopyToArray
Debug.Print "Test Point"
    For acol = LBound(CopyToArray, 2) To UBound(CopyToArray, 2)
        For arow = LBound(CopyToArray, 1) To UBound(CopyToArray, 1)
            Debug.Print c; VarType(c); TypeName(c)
        Next arow
    Next acol


Application.SendKeys "^g ^a {DEL}"

End Sub

Sub PracV2()
    Dim ReturnArray  As Variant
    ReturnArray = Range("A2:L7").Value
'array (rows, columns)
'Debug.Print "Test Point"
'Application.SendKeys "^g ^a {DEL}"
    For arow = LBound(ReturnArray, 1) To UBound(ReturnArray, 1)
        For acol = LBound(ReturnArray, 2) To UBound(ReturnArray, 2)
            'Debug.Print ReturnArray(aCol, aRow)
            If VarType(ReturnArray(arow, acol)) <> 0 Then
                Debug.Print ReturnArray(arow, acol); VarType(ReturnArray(arow, acol)); TypeName(ReturnArray(arow, acol))
            End If
        Next acol
    Debug.Print "-----------------------"
    Next arow
'Application.SendKeys "^g ^a {DEL}"
End Sub
Function CopyToArray() As Variant
    Dim ReturnArray  As Variant
    ReturnArray = Range("A2:L7").Value
    CopyToArray = ReturnArray
End Function




Attribute VB_Name = "Prac"
Sub Prac()
    Dim ReturnArray  As Variant
    ReturnArray = Range("A2:L7").Value
'array (rows, columns)
'Debug.Print "Test Point"
'Application.SendKeys "^g ^a {DEL}"


    For arow = LBound(ReturnArray, 1) To UBound(ReturnArray, 1)
        For acol = LBound(ReturnArray, 2) To UBound(ReturnArray, 2)
            'Debug.Print ReturnArray(aCol, aRow)
            If VarType(ReturnArray(arow, acol)) <> 0 Then
                'Debug.Print ReturnArray(arow, acol); VarType(ReturnArray(arow, acol)); TypeName(ReturnArray(arow, acol))
                'var1 = ReturnArray(arow, acol)
                IsRomNum (ReturnArray(arow, acol))
                
                If IsFoundInDictionary(ReadValues.ReadAllRows()(1).Columns, ReturnArray(arow, acol)) = True Then
                    Debug.Print "Is Row Info Object"
                End If
                
            End If
        Next acol
    Debug.Print "-----------------------"
    Next arow
    
    
    
'Application.SendKeys "^g ^a {DEL}"
End Sub
Function CopyToArray() As Variant
    Dim ReturnArray  As Variant
    CopyToArray = Range("A2:L10").Value
    'ReturnArray = Range("A2:L7").Value
    'CopyToArray = ReturnArray
End Function

Sub TestArray()
    Dim ar As Variant
    ar = CopyToArray()
    Debug.Print "Rows: " & UBound(ar, 1)
    Debug.Print "Cols: " & UBound(ar, 2)
'    Debug.Print ar(1, 22)
'    Debug.Print ar(1, 2)
'    Debug.Print ar(1, 3)
End Sub



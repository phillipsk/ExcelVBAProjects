Attribute VB_Name = "Prac_isObject"

Sub PracIsObjOps()

'Public CollOps As New Collection
'
'Public Eq As String
'Public Add As String
'Public NoEQ As String
'Public Subtract As String
'Public Pos As String
'Public Neg As String

Dim OpsXYZ As New Collection
Dim ops As New Operator



End Sub


Sub PracIsObj()

' Test if variables var1 and var2 represent Object variable types.
Dim var1 As Object
Dim var2, var3, var4
Dim isObj1 As Boolean
Dim isObj2 As Boolean
Dim isObj3 As Boolean
Dim isObj4 As Boolean

isObj1 = IsObject(var1)
' The variable isObj1 is now equal to True.
isObj2 = IsObject(var2)
' The variable isObj2 is now equal to False.
Set var3 = Range("A1")
isObj3 = IsObject(var3)
' The variable isObj3 is now equal to True.
Set var4 = Range("A1")
Set var4 = Nothing
isObj4 = IsObject(var4)
' The variable isObj4 is now equal to True.
'
'The above examples show that the isObject function returns:
'
'    True for variables that are defined as Object types, even if they are not initialized;
'    False for Variants that are not initialized;
'    True for Variants that have objects assigned to them;
'    True for Variants that have previously had objects assigned to them (even if they are now set to Nothing).


End Sub

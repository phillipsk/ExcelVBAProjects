VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Columns As New Dictionary

'Public Id As Integer
'Public Name As String
'Public PriceL As Double
'Public PriceB As Double
'Public FX As Double
'Public Cashier As Integer
'Public WT As String
'Public Organic As String
'Public Vendor As String
'Public Country As String
'Public Store As String
'
Public Category As String
Public RowNumber As String
'
'Public TESTPrice As Double
    ' ========================================================
    ' These Fields are not on the Worksheet
    ' ========================================================
Public amtClose As Double
Public amtOpen As Double

Public totClose As Double
Public totBase As Double


'
'
'op = Operators.Eq
'Debug.Print op
'1
'If op = Operators.Add Then MsgBox ("match")
'var1 = "StoreZZ"
'

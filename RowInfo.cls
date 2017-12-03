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

Public Category As String

Public Id As Integer
Public Name As String
Public Price As Double
Public Cashier As Integer
Public Orangic As String
Public USA As String
Public Store As String



    ' ========================================================
    ' These Fields are not on the Worksheet
    ' ========================================================
Public amtClose As Double
Public amtOpen As Double

Public totClose As Double
Public totBase As Double



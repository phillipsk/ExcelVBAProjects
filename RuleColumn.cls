VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RuleColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Name As String
Public Operator As String
Public value As String
Public ColType As ColumnType
Public Link As String


Enum ColumnType
    [String] = 1
    [Integer] = 2
    [Double] = 3
End Enum


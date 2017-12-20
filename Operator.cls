VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Operator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public CollOps As New Collection

Enum Operators
    Eq = 1
    Gt = 2
    lt = 3
    GtEq = 4
    LtEq = 5
    NoEQ = 6
    Add = 7
    Subtract = 8
    Pos = 9
    Neg = 10
End Enum



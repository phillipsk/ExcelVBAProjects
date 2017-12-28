VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public RuleColumns As New Collection
Public Category As String
Public RuleID As String

Public Total As String

'Evaluating Row upon Rule set if Row is True
Public Success As Boolean

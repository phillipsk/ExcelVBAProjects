VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RuleEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public rules As New Collection
Public Rows As New Collection


Public Sub Apply()

    Dim row As RowInfo
    Dim rule As rule
    For Each row In Rows ' goes through all rows
        ' Test all rules against selected row
        Call ResetRules
        For Each rule In rules
            TestRule row, rule ' takes 1 row and apply 1 rule to it
        Next
    Next
    
End Sub


Sub TestRule(row As RowInfo, rule As rule)
    Dim ruleCol As RuleColumn
    Dim FoundMatch As Boolean: FoundMatch = False ' match in overall rule
    Dim NewMatch As Boolean: NewMatch = False ' match in current rule column
    Dim Link As String: Link = "" ' Link with next condition in the rule
    Dim SkipFor As Boolean
    Dim rule1value As Boolean
    Dim rule2value As Boolean
    
    'Debug.Print row.Columns("Name")
    Dim ruleCounter As Integer
    ruleCounter = 0
    For Each ruleCol In rule.RuleColumns
        'Debug.Print row.RowNumber
        ruleCounter = ruleCounter + 1
        rule1value = False
        rule2value = False
        
        Dim strRowVal As String
        Dim numRowVal As Double
        Dim strRuleVal As String
        Dim numRuleVal As Double
        Dim IsNum As Boolean
        If IsNumeric(ruleCol.Value) Then
            numRowVal = Val(row.Columns(ruleCol.Name))
            numRuleVal = Val(ruleCol.Value)
            IsNum = True
        Else
            strRowVal = row.Columns(ruleCol.Name)
            strRuleVal = ruleCol.Value
            IsNum = False
        End If
        
        If SkipFor Then GoTo ContinueFor
        Select Case ruleCol.Operator
            Case Operators.Eq
                If IsNum Then NewMatch = (numRowVal = numRuleVal) Else NewMatch = (strRowVal = strRuleVal)
            Case Operators.Gt
                If IsNum Then NewMatch = (numRowVal > numRuleVal) Else NewMatch = (strRowVal > strRuleVal)
            Case Operators.Lt
                If IsNum Then NewMatch = (numRowVal < numRuleVal) Else NewMatch = (strRowVal < strRuleVal)
            Case Operators.NoEQ
                If IsNum Then NewMatch = (numRowVal <> numRuleVal) Else NewMatch = (strRowVal <> strRuleVal)
            Case Operators.GtEq
                If IsNum Then NewMatch = (numRowVal >= numRuleVal) Else NewMatch = (strRowVal >= strRuleVal)
            Case Operators.LtEq
                If IsNum Then NewMatch = (numRowVal <= numRuleVal) Else NewMatch = (strRowVal <= strRuleVal)
            Case Operators.And
                ' Get Success value of given rule
                ' E.g. RuleI = True And RuleII = False
                
                rule1value = IsRuleSuccess(ruleCol.Name)
                rule2value = IsRuleSuccess(ruleCol.Value)
                NewMatch = rule1value And rule2value
                
            Case Operators.Or
                
                rule1value = IsRuleSuccess(ruleCol.Name)
                rule2value = IsRuleSuccess(ruleCol.Value)
                NewMatch = rule1value Or rule2value
        
        End Select
        
        ' Evaluate results
        If Link = "AND" Then
            If FoundMatch And NewMatch Then
                row.Category = rule.Category
                rule.Success = True
            End If
        ElseIf Link = "OR" Then
            If FoundMatch Or NewMatch Then
                row.Category = rule.Category
                rule.Success = True
            End If
        Else
            If NewMatch Then
                row.Category = rule.Category
                rule.Success = True
            End If
        End If
        
        FoundMatch = NewMatch
        Link = ruleCol.Link ' Link for next rule column in for loop.
ContinueFor:
    Next
    

End Sub
Function IsRuleSuccess(ruleName As String) As Boolean

    Dim r As rule
    Dim result As Boolean
    For Each r In rules
        If ruleName = "Rule" & r.RuleID Then
            result = r.Success
            If result Then
                'Debug.Print result
            End If
        End If
    Next
    IsRuleSuccess = result
    
End Function

Sub ResetRules()
    Dim r As rule
    For Each r In rules
        r.Success = False
    Next
End Sub


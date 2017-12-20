VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RuleEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Rules As New Collection
Public Rows As New Collection


Public Sub Apply()



    Dim row As RowInfo
    Dim rule As rule
    
    For Each row In Rows ' goes through all rows
        
        ' Test all rules against selected row
        For Each rule In Rules
            TestRule row, rule ' takes 1 row and apply 1 rule to it
        Next
        
    Next
    
End Sub


Sub TestRule(row As RowInfo, rule As rule)
    Dim ruleCol As RuleColumn
    Dim FoundMatch As Boolean
    FoundMatch = False
    Dim SkipFor As Boolean
    
    'Debug.Print row.Columns("Name")
    Dim ruleCounter As Integer
    ruleCounter = 0
    For Each ruleCol In rule.RuleColumns
        Debug.Print row.RowNumber
        ruleCounter = ruleCounter + 1
        
        If SkipFor Then GoTo ContinueFor
        Select Case ruleCol.Operator
            Case Operators.Eq
                'Debug.Print ruleCol.Header & "=" & row.Columns(ruleCol.Header)
                   If row.Columns(ruleCol.Header) = "Apples" Then
                    Debug.Print "Apples"
                   End If
                   
                If row.Columns(ruleCol.Header) = ruleCol.Value Then
                    

                    
                    If ruleCol.Link = "AND" Then
                    
                    Else

                        FoundMatch = True
                    End If
                   
                Else
                    FoundMatch = False
                    SkipFor = True
                End If
        
        End Select
        
        If FoundMatch Then
            'Debug.Print "Match!"
            row.Category = rule.Category
        Else
            'row.Category = ""
        End If
        
ContinueFor:
    Next
    

End Sub


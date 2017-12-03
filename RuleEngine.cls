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
    
    For Each row In Rows

        For Each rule In Rules
            TestRule row, rule
        Next
        
        Debug.Print row.Columns("Id") & "=" & row.Category
    Next
    
End Sub


Sub TestRule(row As RowInfo, rule As rule)
    Dim ruleCol As RuleColumn
    Dim FoundMatch As Boolean
    FoundMatch = False
    
    'Debug.Print row.Columns("Name")
    
    For Each ruleCol In rule.RuleColumns
    
        Select Case ruleCol.Operator
            Case "="
                'Debug.Print ruleCol.Name & "=" & row.Columns(ruleCol.Name)
                If row.Columns(ruleCol.Name) = ruleCol.value Then
                    ' TODO: Refactor and test in all scenarios.
                    If ruleCol.Link = "AND" Then
                        'FoundMatch = FoundMatch And True
                        If row.Columns(ruleCol.Name) = ruleCol.value Then
                            Debug.Print Test
                        End If
                    Else
                        FoundMatch = True
                    End If
                   
                Else
                    FoundMatch = False
                End If
        
        End Select
        
        If FoundMatch Then
            'Debug.Print "Match!"
            row.Category = rule.Category
        End If
    Next
    

End Sub


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
        
'        For Each rule In Rules
'            TestRule row, rule, category
'        Next
        
        
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
        'Debug.Print row.RowNumber
        ruleCounter = ruleCounter + 1
        
        If SkipFor Then GoTo ContinueFor
        Select Case ruleCol.Operator
            Case Operators.Eq
                'Debug.Print ruleCol.Name & "=" & row.Columns(ruleCol.Name)
                If IsNumeric(ruleCol.Value) Then
                    If Val(row.Columns(ruleCol.Name)) = Val(ruleCol.Value) Then
                        
                        If ruleCol.Link <> "AND" Then
                            FoundMatch = True
                        End If
                       
                    Else
                        FoundMatch = False
                        SkipFor = True
                    End If
                Else
                    'If Str(row.Columns(ruleCol.Name)) = Str(ruleCol.Value) Then
                    If row.Columns(ruleCol.Name) = ruleCol.Value Then
                        If ruleCol.Link <> "AND" Then
                            FoundMatch = True
                        End If
                       
                    Else
                        FoundMatch = False
                        SkipFor = True
                    End If
                End If
                
            Case Operators.Gt
                   
'                If rule.RuleID = "II" Then
'                    Debug.Print
'                End If
                
                
                If IsNumeric(ruleCol.Value) Then
                    If Val(row.Columns(ruleCol.Name)) > Val(ruleCol.Value) Then

                        If ruleCol.Link = "AND" Then
                        
                        Else
    
                            FoundMatch = True
                        End If
                       
                    Else
                        FoundMatch = False
                        SkipFor = True
                    End If
                Else
                    If Str(row.Columns(ruleCol.Name)) > Str(ruleCol.Value) Then
                        
                        If ruleCol.Link = "AND" Then
                        
                        Else
    
                            FoundMatch = True
                        End If
                       
                    Else
                        FoundMatch = False
                        SkipFor = True
                    End If
                End If
                            

            Case Operators.Lt
                If IsNumeric(ruleCol.Value) Then
                    If Val(row.Columns(ruleCol.Name)) < Val(ruleCol.Value) Then
                        
                        If ruleCol.Link = "AND" Then
                        
                        Else
    
                            FoundMatch = True
                        End If
                       
                    Else
                        FoundMatch = False
                        SkipFor = True
                    End If
                Else
                    If Str(row.Columns(ruleCol.Name)) < Str(ruleCol.Value) Then
                        
                        If ruleCol.Link = "AND" Then
                        
                        Else
    
                            FoundMatch = True
                        End If
                       
                    Else
                        FoundMatch = False
                        SkipFor = True
                    End If
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


Attribute VB_Name = "ReadValues"
Option Explicit

Sub ReadValues()


    Dim r As Range
    Set r = Range(Range("A2"), Range("A2").End(xlDown))
    
    Dim c As Range
    
    For Each c In r
        Debug.Print c & " - " & c.offset(0, 1)
    Next
    

End Sub


Function ReadAllRows() As Collection

    Dim Rows As New Collection
    
    
    ' Worksheets("Sheet1").range("Id")
    Dim w1 As Worksheet
    Set w1 = Worksheets("Groceries")
    
    ' data range
    Dim r As Range
    Set r = w1.Range(w1.Range("A2"), w1.Range("A2").End(xlDown))
    
    Dim colNames As Range
    Set colNames = Range(w1.Range("A1"), w1.Range("A1").End(xlToRight))
    
    
    
    
    Dim c As Range
    Dim ColCell As Range
    Dim row_info As RowInfo
    
    For Each c In r
    
        Set row_info = New RowInfo
        
        Dim offset As Integer
        offset = 0
        
        For Each ColCell In colNames.Columns
            row_info.Columns.Add ColCell.Value, c.offset(0, offset)
            'Debug.Print offset & ". " & ColCell & " - " & c.offset(0, offset)
            offset = offset + 1
            
        Next
        row_info.RowNumber = c.row
        ' fill the row info fields with values from Columns dictionary.
        Call PopulateRowInfoMembers(row_info)
        Rows.Add row_info
        
        'Debug.Print "row_info: " & row_info.Name
        'Debug.Print c & " - " & c.offset(0, 1)
        
    Next
    
    Set ReadAllRows = Rows
    
End Function



Sub ReadCols()
    Dim r As Range
    Set r = Range(Range("A1"), Range("A1").End(xlToRight))
    
    Dim c As Range
    
    For Each c In r
        'Debug.Print c
    Next
    Debug.Print r.Columns.Count
End Sub


Sub TestRows()

    Dim Rows As Collection
    Set Rows = ReadAllRows()
    
    Dim r As RowInfo
    
    For Each r In Rows
    'Debug.Print "Id: " & r.Columns("Name")
        Dim key As Variant
        For Each key In r.Columns.Keys
            Debug.Print key & ": " & r.Columns(key)
        Next
        
        Debug.Print "-----------------------------"
        
'
'        r.Id = r.Columns("Id")
'        r.Name = r.Columns("Name")
'        r.Price = r.Columns("Price")
'        r.Cashier = r.Columns("Cashier")
'        r.Orangic = r.Columns("Orangic")
'        r.USA = r.Columns("USA")
'        r.Store = r.Columns("Store")
        
    Next

End Sub



Function ReadAllRules() As Collection
    Dim Rules As New Collection
    
    ' read rules from Rules worksheet
    Dim wr As Worksheet
    Set wr = Worksheets("Rules")
    
    Dim r As Range
    ' FIXIT: Fix when there is only one rule. At the moment it selects all rows.
    Set r = wr.Range(wr.Range("A2"), wr.Range("A2").End(xlDown))
    
    Dim c As Range
    Dim rule As New rule
    Dim ruleCol As RuleColumn
    'Debug.Print r.rows.Count
    
    '''' 6 - x??
    '''''' 1 or 0
    
    For Each c In r
    'Debug.Print c.Address
    'Debug.Print c.offset(0, 0)
    
        If c.offset(0, 1) <> "" Then
        'If c <> "" Then Debug.Print "c is not empty"
            Set rule = New rule
            
            Dim OffsetCounter As Integer
            OffsetCounter = 0
            While c.offset(0, 4 + OffsetCounter) <> ""
                Set ruleCol = New RuleColumn
                'ruleCol.ReturnID = c.offset(0, 0 + OffsetCounter)
                ruleCol.Name = c.offset(0, 1 + OffsetCounter)
                ruleCol.Operator = GetOperatorValue(c.offset(0, 2 + OffsetCounter))
                '''Without Enum ruleCol.Operator = c.offset(0, 2 + OffsetCounter)
                ruleCol.Value = c.offset(0, 3 + OffsetCounter)
                ruleCol.Link = c.offset(0, 4 + OffsetCounter)
                'ruleCol.ReturnValue = c.offset(0, 5 + OffsetCounter)
                rule.RuleColumns.Add ruleCol
                
                OffsetCounter = OffsetCounter + 4
            Wend
            
            ' rule category can be first column for simplicity.
            rule.Category = c.offset(0, 1 + OffsetCounter) 'c.End(xlToRight)
            rule.RuleID = c.End(xlToLeft)
            rule.Category = rule.Category + rule.RuleID
            'Debug.Print c
            Rules.Add rule
        End If
    Next
    
    Set ReadAllRules = Rules
End Function


' It populates the fields of RowInfo object using Columns dictionary.
Sub PopulateRowInfoMembers(ByRef ri As RowInfo)
    ri.Id = ri.Columns("Id")
    ri.Name = ri.Columns("Name")
    ri.PriceL = ri.Columns("PriceL")
    ri.PriceB = ri.Columns("PriceB")
    ri.FX = ri.Columns("FX")
    ri.Cashier = ri.Columns("Cashier")
    ri.WT = ri.Columns("WT")
    ri.Organic = ri.Columns("Organic")
    ri.Vendor = ri.Columns("Vendor")
    ri.Country = ri.Columns("Country")
    ri.Store = ri.Columns("Store")
End Sub

'' It populates the fields of RowInfo object using Columns dictionary.
'Function PopulateRowInfoMembers(ByVal ri As RowInfo) As RowInfo
'    ri.Id = ri.Columns("Id")
'    ri.Name = ri.Columns("Name")
'    ri.PriceL = ri.Columns("PriceL")
'    ri.PriceB = ri.Columns("PriceB")
'    ri.FX = ri.Columns("FX")
'    ri.Cashier = ri.Columns("Cashier")
'    ri.WT = ri.Columns("WT")
'    ri.Organic = ri.Columns("Organic")
'    ri.Vendor = ri.Columns("Vendor")
'    ri.Country = ri.Columns("Country")
'    ri.Store = ri.Columns("Store")
'    Set PopulateRowInfoMembers = ri
'End Function
Sub example_Populate()




End Sub

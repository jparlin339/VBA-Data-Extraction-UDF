Attribute VB_Name = "VBAIndexMatch"
Function VBALookup(Sheet As String, ID As Range, IDColName As String, Var As Range, VarRowNum As Integer) As Variant

'UDF designed to triangulate a data point from a separate sheet which contains an external link to a .csv or .txt delimited dataset
'This function works in the following way: INDEX + MATCH variable column + MATCH ID column + MATCH ID within ID column
'Especially useful when considering a self-updating external connection imported into Excel... the function will always find the right value if it exists
'Function will return a #VALUE! error if data does not exist, and will return a "" (blank) value if data value blank
'#VALUE! error may persist if multiple workbooks are open simultaneously
'EXACT MATCH function only; cannot handle duplicates. Must be a proper dataset for function to work properly.
    
    'Determine sheet to retrieve data point given NAME of said sheet (no need for ! to denote sheet reference here, simply need name e.g. DATA)
        Dim Datasheet As Worksheet
        Set Datasheet = ActiveWorkbook.Sheets(Sheet)
        
    'Determine max range and start/end ranges to search through
    LastColumn = Datasheet.UsedRange.Columns.Count
    LastRow = Datasheet.UsedRange.Rows.Count
    ColRange = Range(Datasheet.Cells(VarRowNum, 1), Datasheet.Cells(VarRowNum, LastColumn))
    MaxRange = Range(Cells(1, 1), Cells(LastRow, LastColumn)).Address
    
    'Determine the lookup variable/table header name to search (given cell reference)
    VarColName = Var.Value
    'Determine the dataset record identifier to search (given cell reference)
    Subject = ID.Value
            
    'Calculate column number for lookup
    IndexColNum = WorksheetFunction.Match(VarColName, ColRange, 0)
        'Here we are determining the position of the 'variable' column in question given the specified NAME of said column.. can be anywhere in the dataset
        IDColNum = WorksheetFunction.Match(IDColName, ColRange, 0)
            'Here we are determining the position of the 'identifier'/'reference number' column given the specified NAME of said column.. can be anywhere in the dataset
            Set IDStart = Datasheet.Cells(1, IDColNum)
            Set IDEnd = Datasheet.Cells(LastRow, IDColNum)
                IDRange = IDStart.Address & ":" & IDEnd.Address
                'This is the range of all possible record 'identifiers'/'reference numbers'
    
    'Calculate row number for lookup
    IndexRowNum = WorksheetFunction.Match(Subject, Datasheet.Range(IDRange), 0)
    
    'Compute result using INDEX to triangulate a data point based on CALCULATED row and column numbers within the max range of the dataset
    Result = WorksheetFunction.Index(Datasheet.Range(MaxRange), IndexRowNum, IndexColNum)
    
    'Additional handling of blank values: in the event a data point does not exist, function returns Excel null text string "" and NOT '0'
    'Especially useful for research-based data that typically codes zeroes
        If Result = "" Then
            VBALookup = ""
        Else
            VBALookup = Result
        End If

End Function

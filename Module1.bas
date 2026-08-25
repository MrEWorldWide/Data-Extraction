Attribute VB_Name = "Module1"
Sub LoadDatFile()
Attribute LoadDatFile.VB_Description = "loads the dat file and sorts the information"
Attribute LoadDatFile.VB_ProcData.VB_Invoke_Func = " \n14"
'
' LoadDatFile Macro
' loads the dat file and sorts the information
'

'
Worksheets("Data").Activate

Cells.Clear
   ' With ActiveSheet.QueryTables.Add(Connection:= _
        '"TEXT;C:\Company\Mile high Missoula\Accounting Current\A2NC192960191_attLog.dat", _
        'Destination:=Range("$A$1"))
            With ActiveSheet.QueryTables.Add(Connection:= _
            "TEXT;" & ThisWorkbook.Path & "\A2NC192960191_attLog.dat", _
            Destination:=Range("$A$1"))

        .Name = "A2NC192960191_attLog"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
End Sub
Sub MakeBox()
Attribute MakeBox.VB_ProcData.VB_Invoke_Func = " \n14"
'reset report
Worksheets("Conversion").Activate
    Cells.Clear
    
    Dim DataColumn As Integer
    Dim ReportColumn As Integer
    Dim DataRow As Range
    Dim TotalDataRows As Integer
    Dim RowKeeper As Integer
    
    ' Set TotalDataRows = number of rows of data on Data sheet.
    TotalDataRows = Worksheets("Data").Range("A1").End(xlDown).Row
    'start at column 1
    DataColumn = 1
    RowKeeper = 1
    Dim Emp1, Emp2, Emp3, emp4, Emp5, Emp6, Emp7, Emp8, Emp9, Emp10, Emp11 As String
    Dim LocationName As String
    Dim DateConvert, StartDate, EndDate As Date
    
    Emp1 = Worksheets("Variables").Range("B2").Value
    Emp2 = Worksheets("Variables").Range("B3").Value
    Emp3 = Worksheets("Variables").Range("B4").Value
    emp4 = Worksheets("Variables").Range("B5").Value
    Emp5 = Worksheets("Variables").Range("B6").Value
    Emp6 = Worksheets("Variables").Range("B7").Value
    Emp7 = Worksheets("Variables").Range("B8").Value
    Emp8 = Worksheets("Variables").Range("B9").Value
    Emp9 = Worksheets("Variables").Range("B10").Value
    Emp10 = Worksheets("Variables").Range("B11").Value
    Emp11 = Worksheets("Variables").Range("B12").Value
    
    
    'convert all names
    Do While DataColumn <= TotalDataRows
        
        'check for date range and filter out results
        DateConvert = DateValue(Worksheets("Data").Cells(DataColumn, 2).Value)
        StartDate = DateValue(Worksheets("Variables").Range("F2").Value)
        EndDate = DateValue(Worksheets("Variables").Range("F3").Value)
        
        If DateConvert < StartDate Then GoTo SkipRow
        If DateConvert > EndDate Then GoTo SkipRow

        'convert Emp ID from number to name
        If Worksheets("Data").Cells(DataColumn, 1).Value = 1 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp1
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 2 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp2
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 3 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp3
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 4 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = emp4
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 5 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp5
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 6 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp6
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 7 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp7
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 8 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp8
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 9 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp9
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 10 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp10
        End If
        If Worksheets("Data").Cells(DataColumn, 1).Value = 11 Then
            Worksheets("Conversion").Cells(RowKeeper, 1).Value = Emp11
        End If
        
        
        'convert date and time into schedule
        DateConvert = Worksheets("Data").Cells(DataColumn, 2).Value
        'split of date

        Worksheets("Conversion").Cells(RowKeeper, 2).Value = DateValue(DateConvert)
        'split off time

        Worksheets("Conversion").Cells(RowKeeper, 3).Value = TimeValue(DateConvert)
        
        
        'convert device into location
        If Worksheets("Data").Cells(DataColumn, 3).Value = "A2NC192960191" Then
            Worksheets("Conversion").Cells(RowKeeper, 4).Value = "725 Ronan St #E"
        End If
        
        'convert clock in and out into readable format
        If Worksheets("Data").Cells(DataColumn, 4).Value = 0 Then
            Worksheets("Data").Cells(DataColumn, 4).Value = "In"
        ElseIf Worksheets("Data").Cells(DataColumn, 4).Value = 1 Then
            Worksheets("Data").Cells(DataColumn, 4).Value = "Out"
        End If
        
        Worksheets("Conversion").Cells(RowKeeper, 5).Value = Worksheets("Data").Cells(DataColumn, 4).Value
        RowKeeper = RowKeeper + 1
SkipRow:

        DataColumn = DataColumn + 1
        
    Loop
    'reset counter
        DataColumn = 0

End Sub

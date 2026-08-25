VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
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

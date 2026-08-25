VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ResetData()

'
' ResetReport
'

Worksheets("Finished Report").Activate
    Worksheets.Select
    Cells.Clear

    Range("A1").Value = "Emp ID"
    Range("B1").Value = "Date"
    Range("C1").Value = "Device"
    Range("D1").Value = "Clock out"
    Range("E1").Value = "Clock in"
    Range("F1").Value = "Accumulated hours"
    Range("G1").Value = "Breaks"
    Range("H1").Value = "Total Hours"
    
    
'load finished report
    Dim DataColumn As Integer
    Dim ReportColumn As Integer
    Dim DataRow As Range
    Dim TotalDataRows As Range
    
    ' Set TotalDataRows = number of rows of data on Data sheet.
    TotalDataRows = Worksheets("Data").Range("A2", Range("A2").End(xlDown)).Rows.Count
    'start at column 1
    DataColumn = 1
    
    For Each DataRow In TotalDataRows
        For ReportColumn = 0 To 3
            Worksheet("Finished Report").Range(i).Offset(i, j) = cell.Offset(0, j).Value
        Next ReportColumn
        DataColumn = DataColumn + 1
    Next DataRow
        


End Sub

Private Sub CommandButton1_Click()
MakeBox
End Sub

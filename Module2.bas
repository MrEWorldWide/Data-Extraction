Attribute VB_Name = "Module2"
Sub FinalSort()
Attribute FinalSort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FinalSort Macro
'
'reset report

   'On Error Resume Next
   
Worksheets("Final").Activate
    Cells.Clear
    
    Dim DataColumn As Long
    Dim DataRow As Long
    Dim TotalDataRows As Long
    Dim CurrentName As String
    Dim EmpID As Long
    Dim RowKeeper As Long
    Dim DateConvert As Date
    Dim ClockIn As Boolean
    Dim SkipRow As Boolean
    ClockIn = False
    RowKeeper = 1
    'owner is emp 1
    
    EmpID = 2
    'set the employ name to check
    CurrentName = Worksheets("Variables").Cells(EmpID, 2).Value
    
    ' Set TotalDataRows = number of rows of data on Data sheet.
    TotalDataRows = Worksheets("Conversion").Range("A1").End(xlDown).Row
    DataColumn = 1

    'filter by names, then date, then clock in and clock out times
    'rows 2 through 11
    For EmpID = 2 To 12
    'set employee name to check
    CurrentName = Worksheets("Variables").Cells(EmpID, 2).Value
    
    'loop through all rows on the conversion sheet
        Do While DataColumn <= TotalDataRows
        
            'exit the data transfer if the employee ID is excluded
            SkipRow = False
            If Worksheets("Variables").Cells(EmpID, 3).Value = "x" Then SkipRow = True
            If SkipRow = False Then
            If Worksheets("Conversion").Cells(DataColumn, 1).Value = CurrentName Then

            'convert date and time into schedule
            DateConvert = Worksheets("Conversion").Cells(DataColumn, 2).Value
            
            'split of date
            Worksheets("Final").Cells(RowKeeper, 2).Value = DateValue(DateConvert)
            
            'split off time
            Worksheets("Final").Cells(RowKeeper, 3).Value = TimeValue(DateConvert)
            
            
            Worksheets("Final").Cells(RowKeeper, 1).Value = Worksheets("Conversion").Cells(DataColumn, 1).Value
            Worksheets("Final").Cells(RowKeeper, 2).Value = Worksheets("Conversion").Cells(DataColumn, 2).Value
            Worksheets("Final").Cells(RowKeeper, 3).Value = Worksheets("Conversion").Cells(DataColumn, 3).Value
            Worksheets("Final").Cells(RowKeeper, 4).Value = Worksheets("Conversion").Cells(DataColumn, 4).Value


            'check the in and out status and highlight the row for easy viewing
            If Worksheets("Conversion").Cells(DataColumn, 5).Value = "In" Then
            
                
                With Worksheets("Final").Range(Cells(RowKeeper, 1), Cells(RowKeeper, 10)).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    'this controls the color
                    
                    If ClockIn = True Then
                        'set to red if failed to clock in properly
                        .Color = RGB(255, 185, 185)
                        .TintAndShade = 0
                        'make a note to the right to explain what the error is
                        Worksheets("Final").Cells(RowKeeper, 7).Value = "Failed to Clock out!!!!!!!!!"
                        
                        'set borders around the rows with errors
                        Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlDiagonalDown).LineStyle = xlNone
                        Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlDiagonalUp).LineStyle = xlNone
                        With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlMedium
                        End With
                        With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlMedium
                        End With
                        With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlMedium
                        End With
                        With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlMedium
                        End With
                        Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlInsideVertical).LineStyle = xlNone
                        Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlInsideHorizontal).LineStyle = xlNone
                        
                    Else
                        'set to green if successful clock in
                        .TintAndShade = 0.399975585192419
                    End If
                    
                    
                    
                        
                    
                    ClockIn = True
                    .PatternTintAndShade = 0
                End With
                
                
            ElseIf Worksheets("Conversion").Cells(DataColumn, 5).Value = "Out" Then
            
                With Worksheets("Final").Range(Cells(RowKeeper, 1), Cells(RowKeeper, 10)).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    'this controls the color
                    If ClockIn = False Then
                        'make a note to the right to explain what the error is
                        Worksheets("Final").Cells(RowKeeper, 7).Value = "Failed to Clock in!!!!!!!!!"
                        'set to red if failed to clock in properly
                        .Color = RGB(255, 185, 185)
                        .TintAndShade = 0
                     
                        With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlMedium
                        End With
                        With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlMedium
                        End With
                        With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlMedium
                        End With
                        With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlMedium
                End With
                    Else
                        'set to green if successful clock in
                        .TintAndShade = 0.9
                    End If
                    ClockIn = False
                    .PatternTintAndShade = 0
                End With
                
                
            End If
                
                'set the clock in and out data
                Worksheets("Final").Cells(RowKeeper, 5).Value = Worksheets("Conversion").Cells(DataColumn, 5).Value
                RowKeeper = RowKeeper + 1
            End If


            End If
            'go to next column
            DataColumn = DataColumn + 1
            
        Loop
        'reset counter
        DataColumn = 1
        
        If ClockIn = True Then
        'highlight the row red if they fail to clock in properly
            With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent3
                'make a note to the right to explain what the error is
                Worksheets("Final").Cells(RowKeeper - 1, 7).Value = "Failed to Clock out!!!!!!!!!"
                
                'RowKeeper = RowKeeper + 1
                .Color = RGB(255, 185, 185)
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With
                
                'set borders around the rows with errors
                Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlDiagonalDown).LineStyle = xlNone
                Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlDiagonalUp).LineStyle = xlNone
                With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlInsideVertical).LineStyle = xlNone
                Worksheets("Final").Range(Cells(RowKeeper - 1, 1), Cells(RowKeeper, 10)).Borders(xlInsideHorizontal).LineStyle = xlNone
                
                
                ClockIn = False
        End If
        
        If RowKeeper > 1 Then
            RowKeeper = RowKeeper + 1
        End If
        
    Next EmpID
    
    
End Sub

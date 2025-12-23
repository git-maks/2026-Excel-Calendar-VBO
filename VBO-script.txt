Sub Generate2026Calendar_v3()
    Dim ws As Worksheet
    Dim yearVal As Integer
    Dim i As Integer, monthVal As Integer
    Dim startDate As Date, endDate As Date, curDate As Date
    Dim startRow As Integer, curRow As Integer, curCol As Integer, startCol As Integer
    Dim headerRange As Range, daysHeaderRange As Range
    Dim dayNames As Variant
    Dim monthColors(1 To 12) As Long
    
    ' Variables for Dark Color Calculation
    Dim baseColor As Long
    Dim r As Long, g As Long, b As Long
    
    ' --- CONFIGURATION ---
    yearVal = 2026
    startCol = 2 ' Start at Column B (Column 2)
    
    ' Define Colors (Converted from HEX to RGB Long)
    monthColors(1) = RGB(52, 125, 178)   ' JAN #347DB2
    monthColors(2) = RGB(57, 186, 198)   ' FEB #39BAC6
    monthColors(3) = RGB(57, 198, 153)   ' MAR #39C699
    monthColors(4) = RGB(59, 206, 98)    ' APR #3BCE62
    monthColors(5) = RGB(132, 206, 80)   ' MAY #84CE50
    monthColors(6) = RGB(229, 211, 108)  ' JUN #E5D36C
    monthColors(7) = RGB(230, 182, 86)   ' JUL #E6B656
    monthColors(8) = RGB(217, 111, 58)   ' AUG #D96F3A
    monthColors(9) = RGB(192, 64, 53)    ' SEP #C04035
    monthColors(10) = RGB(174, 50, 94)   ' OCT #AE325E
    monthColors(11) = RGB(138, 61, 164)  ' NOV #8A3DA4
    monthColors(12) = RGB(74, 84, 191)   ' DEC #4A54BF
    
    ' Day Names (Monday Start) - Title Case
    dayNames = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    
    ' --- FIX: DELETE EXISTING SHEET IF IT EXISTS ---
    Dim sheetName As String
    sheetName = "Calendar " & yearVal
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    ' -----------------------------------------------

    ' Create new sheet
    Set ws = Worksheets.Add
    ws.Name = sheetName
    
    ' --- Set ENTIRE Sheet Background to White ---
    ws.Cells.Interior.Color = vbWhite
    
    ' Set Default Alignment to MIDDLE (Center) for EVERYTHING
    ws.Cells.VerticalAlignment = xlCenter
    
    ' Set Padding/WeekNum Column A width
    ws.Columns(1).ColumnWidth = 3
    ' Set Calendar Columns (B to H) width
    ws.Range(ws.Columns(startCol), ws.Columns(startCol + 6)).ColumnWidth = 23
    
    ' ==========================================================================================
    ' NEW FEATURE: YEAR PROGRESS BAR (Row 2)
    ' ==========================================================================================
    curRow = 2
    Dim progressBarRange As Range
    Set progressBarRange = ws.Range(ws.Cells(curRow, startCol), ws.Cells(curRow, startCol + 6))
    
    With progressBarRange
        .Merge
        ' Insert Formula: Calculates percentage of year passed based on TODAY
        ' Formula logic: (Today - Jan1) / (Dec31 - Jan1)
        .Formula = "=MAX(0, MIN(1, (TODAY() - DATE(" & yearVal & ",1,1)) / (DATE(" & yearVal & ",12,31) - DATE(" & yearVal & ",1,1))))"
        
        ' Formatting
        .Style = "Percent"
        .Font.Name = "Atopos"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 35
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThick
        
        ' Add Data Bar (The visual filler)
        Dim db As Databar
        Set db = .FormatConditions.AddDatabar
        db.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        db.MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
        
        ' --- COLOR CHANGED TO DARK GRAY ---
        db.BarColor.Color = RGB(64, 64, 64) 
        db.BarFillType = xlDataBarFillSolid
    End With
    
    ' Custom Text Label Logic for the Progress Bar
    progressBarRange.NumberFormat = """Year Completion: ""0%"
    
    ' Skip a row for spacing
    curRow = curRow + 2 ' Now we are at Row 4
    
    ' ==========================================================================================
    ' GENERATE MONTHS
    ' ==========================================================================================
    For monthVal = 1 To 12
        startDate = DateSerial(yearVal, monthVal, 1)
        endDate = DateSerial(yearVal, monthVal + 1, 0)
        
        ' 1. Month Header
        Set headerRange = ws.Range(ws.Cells(curRow, startCol), ws.Cells(curRow, startCol + 6))
        With headerRange
            .Merge
            .Value = MonthName(monthVal)
            .Interior.Color = monthColors(monthVal)
            .Font.Color = vbWhite
            .Font.Name = "Atopos"
            .Font.Bold = True
            .Font.Size = 16
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 30
            .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        End With
        
        curRow = curRow + 1
        
        ' 2. Days of Week Header
        Set daysHeaderRange = ws.Range(ws.Cells(curRow, startCol), ws.Cells(curRow, startCol + 6))
        For i = 0 To 6
            With ws.Cells(curRow, startCol + i)
                .Value = dayNames(i)
                .Interior.Color = RGB(242, 242, 242)
                .Font.Name = "Atopos"
                .Font.Bold = True
                .Font.Size = 12
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .RowHeight = 24
                .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
            End With
        Next i
        
        curRow = curRow + 1
        startRow = curRow
        
        ' 3. Fill Dates & Empty Slots
        curDate = startDate
        
        ' --- PRE-PADDING ---
        Dim firstDayWeekday As Integer
        firstDayWeekday = Weekday(startDate, vbMonday)
        
        If firstDayWeekday > 1 Then
            For i = 1 To firstDayWeekday - 1
                curCol = startCol + (i - 1)
                
                Dim emptyBlockRange As Range
                Set emptyBlockRange = ws.Range(ws.Cells(curRow, curCol), ws.Cells(curRow + 3, curCol))
                
                With emptyBlockRange
                    .Interior.Color = vbWhite
                    .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
                    .Cells(1, 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Cells(1, 1).Borders(xlEdgeBottom).Weight = xlThin
                    .Cells(2, 1).Borders(xlEdgeBottom).LineStyle = xlDot
                    .Cells(2, 1).Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
                    .Cells(3, 1).Borders(xlEdgeBottom).LineStyle = xlDot
                    .Cells(3, 1).Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
                End With
            Next i
        End If
        
        ' --- ACTUAL DATES LOOP ---
        Do While Month(curDate) = monthVal
            curCol = startCol + (Weekday(curDate, vbMonday) - 1)
            
            ' --- Date Cell Content ---
            With ws.Cells(curRow, curCol)
                ' CRITICAL UPDATE: Use Actual Date Value
                .Value = curDate
                ' Format it to look like "01.01"
                .NumberFormat = "dd.mm"
                
                .Font.Name = "Atopos"
                .Font.Bold = True
                .Font.Size = 11
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                
                ' Darken the Month Color for Text
                baseColor = monthColors(monthVal)
                r = (baseColor Mod 256) * 0.5
                g = ((baseColor \ 256) Mod 256) * 0.5
                b = ((baseColor \ 65536) Mod 256) * 0.5
                .Font.Color = RGB(r, g, b)
            End With
            
            ' --- Define Block & Styles ---
            Dim blockRange As Range
            Set blockRange = ws.Range(ws.Cells(curRow, curCol), ws.Cells(curRow + 3, curCol))
            
            With blockRange
                .Interior.Color = vbWhite
                .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
                
                With .Cells(1, 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                
                ' Set Note Font
                .Cells(2, 1).Font.Name = "Atopos Narrow"
                .Cells(2, 1).Font.Size = 10
                .Cells(3, 1).Font.Name = "Atopos Narrow"
                .Cells(3, 1).Font.Size = 10
                .Cells(4, 1).Font.Name = "Atopos Narrow"
                .Cells(4, 1).Font.Size = 10
                
                ' Internal Dotted Lines
                .Cells(2, 1).Borders(xlEdgeBottom).LineStyle = xlDot
                .Cells(2, 1).Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
                .Cells(3, 1).Borders(xlEdgeBottom).LineStyle = xlDot
                .Cells(3, 1).Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
            End With
            
            If Weekday(curDate, vbMonday) = 7 Then
                curRow = curRow + 4
            End If
            
            curDate = curDate + 1
        Loop
        
        ' --- POST-PADDING ---
        Dim lastDayWeekday As Integer
        lastDayWeekday = Weekday(curDate - 1, vbMonday)
        
        If lastDayWeekday < 7 Then
            For i = lastDayWeekday + 1 To 7
                curCol = startCol + (i - 1)
                
                Dim postEmptyBlockRange As Range
                Set postEmptyBlockRange = ws.Range(ws.Cells(curRow, curCol), ws.Cells(curRow + 3, curCol))
                
                With postEmptyBlockRange
                    .Interior.Color = vbWhite
                    .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
                    .Cells(1, 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Cells(1, 1).Borders(xlEdgeBottom).Weight = xlThin
                    .Cells(2, 1).Borders(xlEdgeBottom).LineStyle = xlDot
                    .Cells(2, 1).Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
                    .Cells(3, 1).Borders(xlEdgeBottom).LineStyle = xlDot
                    .Cells(3, 1).Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
                End With
            Next i
            curRow = curRow + 4
        End If
        
        curRow = curRow + 2
    Next monthVal
    
    ' ==========================================================================================
    ' NEW FEATURE: AUTOMATIC CURRENT DAY HIGHLIGHTING
    ' ==========================================================================================
    Dim cfRange As Range
    Set cfRange = ws.Range(ws.Columns(startCol), ws.Columns(startCol + 6))
    
    Dim cond As FormatCondition
    Set cond = cfRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TODAY()")
    
    With cond
        .Interior.Color = RGB(192, 0, 0) ' Deep Red Background
        .Font.Color = vbWhite            ' White Text
        .Font.Bold = True                ' Bold Text
    End With
    
    ' Fix layout refresh
    ws.Columns.AutoFit
    ws.Range(ws.Columns(startCol), ws.Columns(startCol + 6)).ColumnWidth = 23
    
    MsgBox "Calendar for " & yearVal & " generated successfully!", vbInformation
End Sub

' =================================
' Rendering Module - VBA Version
' Converted from Google Apps Script
' =================================

Option Explicit

' Render payslip to Excel worksheet
Public Sub RenderPaySlip(employeeName As String, startOfWeek As Date, endOfWeek As Date, _
                        parsedShiftData As Object, summary As Object, weeklyTotal As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim baseHeaders As Variant
    Dim allWageTypes As Collection
    Dim headers As Variant
    Dim rowIndex As Long
    Dim dayKey As Variant
    Dim dayRecords As Collection
    Dim i As Integer
    Dim record As ShiftRecord
    Dim row As Variant
    Dim j As Integer
    Dim wageType As Variant
    Dim entry As Object
    Dim sortedKeys As Collection
    Dim key As Variant
    Dim item As Object
    Dim line As Variant
    
    ' Get or create payslip worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Payslip")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Payslip"
    End If
    
    ' Clear existing content (keep input cells A1, B1)
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow > 1 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
    End If
    
    ' === Payslip Header ===
    ws.Cells(2, 1).Value = "Weekly Payslip"
    ws.Cells(2, 1).Font.Bold = True
    ws.Cells(2, 1).Font.Size = 16
    
    ws.Cells(3, 1).Value = "Name: " & employeeName
    ws.Cells(4, 1).Value = "Week: " & FormatDate(startOfWeek) & " to " & FormatDate(endOfWeek)
    
    ' === Table Headers ===
    baseHeaders = Array("Date", "Day", "Start", "End", "Break", "Total")
    Set allWageTypes = CollectAllWageTypes(parsedShiftData, summary)
    
    ' Combine base headers with wage types
    ReDim headers(0 To UBound(baseHeaders) + allWageTypes.Count)
    For i = 0 To UBound(baseHeaders)
        headers(i) = baseHeaders(i)
    Next i
    
    For i = 1 To allWageTypes.Count
        headers(UBound(baseHeaders) + i) = allWageTypes(i)
    Next i
    
    ws.Cells(6, 1).Value = "Shift Logs"
    ws.Cells(6, 1).Font.Bold = True
    
    ' Write headers
    For i = 0 To UBound(headers)
        ws.Cells(7, i + 1).Value = headers(i)
        ws.Cells(7, i + 1).Font.Bold = True
    Next i
    
    ' === Render Shift Logs ===
    rowIndex = 8
    
    ' Iterate through each day in sorted order
    For Each dayKey In parsedShiftData.Keys
        Set dayRecords = parsedShiftData(dayKey)
        
        For i = 1 To dayRecords.Count
            record = dayRecords(i)
            
            ' Build row data
            ReDim row(0 To UBound(headers))
            row(0) = FormatDate(record.shiftDate)
            row(1) = GetDayName(record.shiftDate)
            row(2) = FormatTime(record.startTime)
            row(3) = FormatTime(record.finishTime)
            row(4) = record.breakHours
            row(5) = RoundToTwo(GetDurationHours(record.startTime, record.finishTime) - record.breakHours)
            
            ' Fill wage categories
            For j = 1 To allWageTypes.Count
                wageType = allWageTypes(j)
                If record.parsedShift.Exists(wageType) Then
                    Set entry = record.parsedShift(wageType)
                    If RoundToTwo(entry("hours")) = 0 Then
                        row(UBound(baseHeaders) + j) = ""
                    Else
                        row(UBound(baseHeaders) + j) = RoundToTwo(entry("hours"))
                    End If
                Else
                    row(UBound(baseHeaders) + j) = ""
                End If
            Next j
            
            ' Write row to worksheet
            For j = 0 To UBound(row)
                ws.Cells(rowIndex, j + 1).Value = row(j)
            Next j
            
            rowIndex = rowIndex + 1
        Next i
    Next dayKey
    
    ' === Summary Block ===
    rowIndex = rowIndex + 2
    ws.Cells(rowIndex, 1).Value = "Summary"
    ws.Cells(rowIndex, 1).Font.Bold = True
    rowIndex = rowIndex + 1
    
    ' Sort summary keys by wage
    Set sortedKeys = SortSummaryKeysByWage(summary)
    
    For i = 1 To sortedKeys.Count
        key = sortedKeys(i)
        Set item = summary(key)
        
        If RoundToTwo(item("total")) > 0 Then
            ws.Cells(rowIndex, 1).Value = key & ": " & RoundToTwo(item("hours")) & " hours"
            ws.Cells(rowIndex, 2).Value = "at $" & RoundToTwo(item("wage"))
            ws.Cells(rowIndex, 3).Value = "$" & RoundToTwo(item("total"))
            rowIndex = rowIndex + 1
        End If
    Next i
    
    ' === Total Pay ===
    ws.Cells(rowIndex, 1).Value = "Total"
    ws.Cells(rowIndex, 1).Font.Bold = True
    ws.Cells(rowIndex, 3).Value = "$" & RoundToTwo(weeklyTotal("total"))
    ws.Cells(rowIndex, 3).Font.Bold = True
    
    ' Auto-fit columns
    ws.Columns.AutoFit
    
    ' Add borders to the table
    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(7, 1), ws.Cells(rowIndex, UBound(headers) + 1))
    tableRange.Borders.LineStyle = xlContinuous
    
    ' Format header row
    ws.Range(ws.Cells(7, 1), ws.Cells(7, UBound(headers) + 1)).Interior.Color = RGB(200, 200, 200)
    
    MsgBox "Payslip rendered successfully!", vbInformation, "Payroll System"
End Sub

' Collect all wage types from parsed shift data
Private Function CollectAllWageTypes(parsedShiftData As Object, summary As Object) As Collection
    Dim typesSet As Object
    Dim dayKey As Variant
    Dim dayRecords As Collection
    Dim i As Integer
    Dim record As ShiftRecord
    Dim key As Variant
    Dim typesCollection As New Collection
    Dim sortedTypes As Collection
    
    Set typesSet = CreateObject("Scripting.Dictionary")
    
    ' Collect all unique wage types
    For Each dayKey In parsedShiftData.Keys
        Set dayRecords = parsedShiftData(dayKey)
        
        For i = 1 To dayRecords.Count
            record = dayRecords(i)
            If Not record.parsedShift Is Nothing Then
                For Each key In record.parsedShift.Keys
                    If record.parsedShift(key)("hours") > 0 Then
                        If Not typesSet.Exists(key) Then
                            typesSet(key) = True
                        End If
                    End If
                Next key
            End If
        Next i
    Next dayKey
    
    ' Convert to collection
    For Each key In typesSet.Keys
        typesCollection.Add key
    Next key
    
    ' Sort by wage value from summary
    Set sortedTypes = SortWageTypesByWage(typesCollection, summary)
    
    Set CollectAllWageTypes = sortedTypes
End Function

' Sort wage types by their wage values
Private Function SortWageTypesByWage(wageTypes As Collection, summary As Object) As Collection
    Dim sortedTypes As New Collection
    Dim wages() As Double
    Dim types() As String
    Dim i As Integer, j As Integer
    Dim temp As Double
    Dim tempStr As String
    Dim wageType As Variant
    
    ' Convert to arrays for sorting
    ReDim wages(1 To wageTypes.Count)
    ReDim types(1 To wageTypes.Count)
    
    For i = 1 To wageTypes.Count
        wageType = wageTypes(i)
        types(i) = wageType
        If summary.Exists(wageType) Then
            wages(i) = summary(wageType)("wage")
        Else
            wages(i) = 0
        End If
    Next i
    
    ' Bubble sort
    For i = 1 To UBound(wages) - 1
        For j = i + 1 To UBound(wages)
            If wages(i) > wages(j) Then
                ' Swap wages
                temp = wages(i)
                wages(i) = wages(j)
                wages(j) = temp
                ' Swap types
                tempStr = types(i)
                types(i) = types(j)
                types(j) = tempStr
            End If
        Next j
    Next i
    
    ' Convert back to collection
    For i = 1 To UBound(types)
        sortedTypes.Add types(i)
    Next i
    
    Set SortWageTypesByWage = sortedTypes
End Function

' Sort summary keys by wage value
Private Function SortSummaryKeysByWage(summary As Object) As Collection
    Dim sortedKeys As New Collection
    Dim wages() As Double
    Dim keys() As String
    Dim i As Integer, j As Integer
    Dim temp As Double
    Dim tempStr As String
    Dim key As Variant
    Dim keyCount As Integer
    
    keyCount = summary.Count
    ReDim wages(1 To keyCount)
    ReDim keys(1 To keyCount)
    
    ' Fill arrays
    i = 1
    For Each key In summary.Keys
        keys(i) = key
        wages(i) = summary(key)("wage")
        i = i + 1
    Next key
    
    ' Bubble sort
    For i = 1 To keyCount - 1
        For j = i + 1 To keyCount
            If wages(i) > wages(j) Then
                ' Swap wages
                temp = wages(i)
                wages(i) = wages(j)
                wages(j) = temp
                ' Swap keys
                tempStr = keys(i)
                keys(i) = keys(j)
                keys(j) = tempStr
            End If
        Next j
    Next i
    
    ' Convert to collection
    For i = 1 To keyCount
        sortedKeys.Add keys(i)
    Next i
    
    Set SortSummaryKeysByWage = sortedKeys
End Function

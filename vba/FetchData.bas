' =================================
' Data Fetching Module - VBA Version
' Converted from Google Apps Script
' =================================

Option Explicit

' Structure to hold sheet data
Public Type SheetData
    headers As Variant
    rows As Variant
End Type

' Structure to hold shift record
Public Type ShiftRecord
    shiftDate As Date
    employeeName As String
    startTime As Date
    finishTime As Date
    breakHours As Double
    parsedShift As Object ' Dictionary for parsed shift data
End Type

' Structure to hold config data
Public Type ConfigData
    openTime As Date
    closeTime As Date
    otWeeklyTimeThreshold As Variant
    otWeeklyThresholdWage As Variant
    otDailyTimeThreshold As Variant
    otDailyThresholdWage As Variant
    wdBaseWage As Double
    wdEarlyOtWage As Double
    wdLateOtWage As Double
    satBaseWage As Double
    satEarlyOtWage As Double
    satLateOtWage As Double
    sunBaseWage As Double
    sunEarlyOtWage As Double
    sunLateOtWage As Double
End Type

' Read data from a worksheet
Public Function ReadSheet(sheetName As String) As SheetData
    Dim ws As Worksheet
    Dim data As SheetData
    Dim lastRow As Long
    Dim lastCol As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Then
        MsgBox "No data found in sheet: " & sheetName
        Exit Function
    End If
    
    ' Get headers (first row)
    data.headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value
    
    ' Get data rows (skip header)
    If lastRow > 1 Then
        data.rows = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
    End If
    
    ReadSheet = data
    Exit Function
    
ErrorHandler:
    MsgBox "Error reading sheet '" & sheetName & "': " & Err.Description
End Function

' Load configuration from Config sheet
Public Function GetConfig() As ConfigData
    Dim sheetData As SheetData
    Dim config As ConfigData
    Dim errors As Collection
    Dim i As Integer
    Dim header As String
    Dim value As Variant
    
    Set errors = New Collection
    sheetData = ReadSheet("Config")
    
    If IsEmpty(sheetData.headers) Then
        MsgBox "Failed to read Config sheet"
        Exit Function
    End If
    
    ' Process each header and its corresponding value
    For i = 1 To UBound(sheetData.headers, 2)
        header = Trim(sheetData.headers(1, i))
        value = sheetData.rows(1, i)
        
        On Error GoTo ConfigError
        
        Select Case header
            Case "OPEN_TIME"
                If Not IsDate(value) Then
                    errors.Add "Invalid Time for " & header
                Else
                    config.openTime = CDate(value)
                End If
                
            Case "CLOSE_TIME"
                If Not IsDate(value) Then
                    errors.Add "Invalid Time for " & header
                Else
                    config.closeTime = CDate(value)
                End If
                
            Case "OT_WEEKLY_TIME_THRESHOLD"
                config.otWeeklyTimeThreshold = Split(CStr(value), ",")
                
            Case "OT_WEEKLY_THRESHOLD_WAGE"
                config.otWeeklyThresholdWage = Split(CStr(value), ",")
                
            Case "OT_DAILY_TIME_THRESHOLD"
                config.otDailyTimeThreshold = Split(CStr(value), ",")
                
            Case "OT_DAILY_THRESHOLD_WAGE"
                config.otDailyThresholdWage = Split(CStr(value), ",")
                
            Case "WD_BASE_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.wdBaseWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
                
            Case "WD_EARLY_OT_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.wdEarlyOtWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
                
            Case "WD_LATE_OT_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.wdLateOtWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
                
            Case "SAT_BASE_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.satBaseWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
                
            Case "SAT_EARLY_OT_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.satEarlyOtWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
                
            Case "SAT_LATE_OT_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.satLateOtWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
                
            Case "SUN_BASE_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.sunBaseWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
                
            Case "SUN_EARLY_OT_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.sunEarlyOtWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
                
            Case "SUN_LATE_OT_WAGE"
                If IsNumeric(value) And value >= 0 Then
                    config.sunLateOtWage = CDbl(value)
                Else
                    errors.Add "Invalid wage value for " & header
                End If
        End Select
        GoTo NextHeader
        
ConfigError:
        errors.Add "Error processing " & header & ": " & Err.Description
        
NextHeader:
        On Error GoTo 0
    Next i
    
    If errors.Count > 0 Then
        RaiseErrors "Critical Config Errors found. Calculation cannot proceed:" & vbCrLf & vbCrLf, errors
    End If
    
    GetConfig = config
End Function

' Validate and get shift data from Shift Entry sheet
Public Function GetValidatedShiftData() As Collection
    Dim sheetData As SheetData
    Dim errors As Collection
    Dim result As Collection
    Dim record As ShiftRecord
    Dim i As Long, j As Integer
    Dim header As String
    Dim value As Variant
    Dim hasError As Boolean
    
    Set errors = New Collection
    Set result = New Collection
    sheetData = ReadSheet("Shift Entry")
    
    If IsEmpty(sheetData.headers) Or IsEmpty(sheetData.rows) Then
        MsgBox "No data found in Shift Entry sheet"
        Set GetValidatedShiftData = result
        Exit Function
    End If
    
    ' Process each row
    For i = 1 To UBound(sheetData.rows, 1)
        hasError = False
        Set record.parsedShift = CreateObject("Scripting.Dictionary")
        
        ' Process each column in the row
        For j = 1 To UBound(sheetData.headers, 2)
            header = Trim(sheetData.headers(1, j))
            value = sheetData.rows(i, j)
            
            On Error GoTo ValidationError
            
            Select Case header
                Case "Date"
                    If Not IsDate(value) Then
                        errors.Add "Row " & (i + 1) & ", Column '" & header & "': Invalid Date"
                        hasError = True
                    Else
                        record.shiftDate = CDate(value)
                    End If
                    
                Case "Name"
                    If IsEmpty(value) Or Trim(CStr(value)) = "" Then
                        errors.Add "Row " & (i + 1) & ", Column '" & header & "': Name is Missing"
                        hasError = True
                    Else
                        record.employeeName = Trim(CStr(value))
                    End If
                    
                Case "Start Time"
                    If Not IsDate(value) Then
                        errors.Add "Row " & (i + 1) & ", Column '" & header & "': Invalid Time"
                        hasError = True
                    Else
                        record.startTime = CDate(value)
                    End If
                    
                Case "Finish Time"
                    If Not IsDate(value) Then
                        errors.Add "Row " & (i + 1) & ", Column '" & header & "': Invalid Time"
                        hasError = True
                    Else
                        record.finishTime = CDate(value)
                    End If
                    
                Case "Break(Hours)"
                    If IsEmpty(value) Or value = "" Then
                        value = 0
                    End If
                    If Not IsNumeric(value) Or CDbl(value) < 0 Then
                        errors.Add "Row " & (i + 1) & ", Column '" & header & "': Invalid Break Value"
                        hasError = True
                    Else
                        record.breakHours = CDbl(value)
                        ' Validate break time doesn't exceed shift duration
                        If record.breakHours > GetDurationHours(record.startTime, record.finishTime) Then
                            errors.Add "Row " & (i + 1) & ", Column '" & header & "': Break Time is larger than Shift Hours"
                            hasError = True
                        End If
                    End If
            End Select
            GoTo NextColumn
            
ValidationError:
            errors.Add "Row " & (i + 1) & ", Column '" & header & "': " & Err.Description
            hasError = True
            
NextColumn:
            On Error GoTo 0
        Next j
        
        If Not hasError Then
            result.Add record
        End If
    Next i
    
    If errors.Count > 0 Then
        RaiseErrors "Some records have validation errors and were excluded:" & vbCrLf & vbCrLf, errors
    End If
    
    Set GetValidatedShiftData = result
End Function

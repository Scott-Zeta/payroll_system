' =================================
' Data Processing Module - VBA Version
' Converted from Google Apps Script
' =================================

Option Explicit

' Sort and group shift data by day of week
Public Function SortAndGroupByDate(data As Collection) As Object
    Dim groupedMap As Object
    Dim sortedArray() As ShiftRecord
    Dim i As Integer, j As Integer
    Dim record As ShiftRecord
    Dim dayKey As Integer
    Dim temp As ShiftRecord
    
    Set groupedMap = CreateObject("Scripting.Dictionary")
    
    ' Convert collection to array for sorting
    ReDim sortedArray(1 To data.Count)
    For i = 1 To data.Count
        sortedArray(i) = data(i)
    Next i
    
    ' Bubble sort by date, then by start time
    For i = 1 To UBound(sortedArray)
        For j = i + 1 To UBound(sortedArray)
            If sortedArray(i).shiftDate > sortedArray(j).shiftDate Or _
               (sortedArray(i).shiftDate = sortedArray(j).shiftDate And _
                sortedArray(i).startTime > sortedArray(j).startTime) Then
                ' Swap records
                temp = sortedArray(i)
                sortedArray(i) = sortedArray(j)
                sortedArray(j) = temp
            End If
        Next j
    Next i
    
    ' Group by day of week (1=Monday, 7=Sunday)
    For i = 1 To UBound(sortedArray)
        record = sortedArray(i)
        dayKey = Weekday(record.shiftDate, vbMonday) ' 1=Monday, 7=Sunday
        
        If Not groupedMap.Exists(dayKey) Then
            Set groupedMap(dayKey) = New Collection
        End If
        
        groupedMap(dayKey).Add record
    Next i
    
    Set SortAndGroupByDate = groupedMap
End Function

' Parse shifts and calculate wages
Public Function ParseShift(groupedShiftMap As Object) As Object
    Dim summary As Object
    Dim weeklyTotal As Double
    Dim dayKey As Variant
    Dim dayRecords As Collection
    Dim i As Integer
    Dim record As ShiftRecord
    Dim parseResult As Object
    Dim breakTime As Double
    Dim duration As Double
    Dim dailyTotal As Double
    Dim weeklyRemain As Double
    Dim dailyRemain As Double
    Dim weeklyOTResult As Object
    Dim dailyOTResult As Object
    Dim regularWorkResult As Object
    Dim key As Variant
    
    Set summary = CreateObject("Scripting.Dictionary")
    weeklyTotal = 0
    
    ' Process each day
    For Each dayKey In groupedShiftMap.Keys
        Set dayRecords = groupedShiftMap(dayKey)
        dailyTotal = 0
        
        Debug.Print "Processing Day: " & dayKey
        
        ' Process each shift record for this day
        For i = 1 To dayRecords.Count
            record = dayRecords(i)
            Set parseResult = CreateObject("Scripting.Dictionary")
            
            breakTime = Application.WorksheetFunction.Max(record.breakHours, 0)
            duration = GetDurationHours(record.startTime, record.finishTime) - breakTime
            weeklyTotal = weeklyTotal + duration
            dailyTotal = dailyTotal + duration
            
            Debug.Print "Weekly Total: " & weeklyTotal & ", Daily Total: " & dailyTotal
            
            ' Parse weekly overtime first
            Set weeklyOTResult = ParseOvertime("Weekly", weeklyTotal, duration, _
                globalConfig.otWeeklyTimeThreshold, globalConfig.otWeeklyThresholdWage, weeklyRemain)
            
            ' Merge weekly OT results
            For Each key In weeklyOTResult.Keys
                If key <> "hoursRemain" Then
                    Set parseResult(key) = weeklyOTResult(key)
                End If
            Next key
            weeklyRemain = weeklyOTResult("hoursRemain")
            
            Debug.Print "Hours remaining after weekly OT: " & weeklyRemain
            
            ' If there are remaining hours, check for daily overtime
            If weeklyRemain > 0 Then
                Set dailyOTResult = ParseOvertime("Daily", weeklyRemain, weeklyRemain, _
                    globalConfig.otDailyTimeThreshold, globalConfig.otDailyThresholdWage, dailyRemain)
                
                ' Merge daily OT results
                For Each key In dailyOTResult.Keys
                    If key <> "hoursRemain" Then
                        Set parseResult(key) = dailyOTResult(key)
                    End If
                Next key
                dailyRemain = dailyOTResult("hoursRemain")
                
                ' If there are still remaining hours, parse as regular work
                If dailyRemain > 0 Then
                    Set regularWorkResult = ParseRegularWork(record.shiftDate, record.startTime, dailyRemain, record.breakHours)
                    
                    ' Merge regular work results
                    For Each key In regularWorkResult.Keys
                        Set parseResult(key) = regularWorkResult(key)
                    Next key
                End If
            End If
            
            ' Store parsed result in record
            Set record.parsedShift = parseResult
            
            ' Update summary
            For Each key In parseResult.Keys
                If Not summary.Exists(key) Then
                    Set summary(key) = CreateObject("Scripting.Dictionary")
                    summary(key)("wage") = parseResult(key)("wage")
                    summary(key)("hours") = 0
                    summary(key)("total") = 0
                End If
                summary(key)("hours") = summary(key)("hours") + parseResult(key)("hours")
                summary(key)("total") = summary(key)("total") + parseResult(key)("total")
            Next key
        Next i
    Next dayKey
    
    Set ParseShift = summary
End Function

' Parse overtime hours based on thresholds
Private Function ParseOvertime(prefix As String, totalHours As Double, duration As Double, _
                              thresholds As Variant, thresholdWages As Variant, ByRef hoursRemain As Double) As Object
    Dim result As Object
    Dim splitArray() As Double
    Dim start As Double
    Dim remaining As Double
    Dim i As Integer
    Dim endThreshold As Double
    Dim rangeStart As Double
    Dim rangeEnd As Double
    Dim hoursInRange As Double
    Dim key As String
    
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Initialize split array
    ReDim splitArray(0 To UBound(thresholds) + 1)
    For i = 0 To UBound(splitArray)
        splitArray(i) = 0
    Next i
    
    start = totalHours - duration
    remaining = duration
    
    Debug.Print "ParseOvertime - " & prefix & ": Total=" & totalHours & ", Duration=" & duration
    
    ' Distribute hours across thresholds
    For i = 0 To UBound(thresholds) + 1
        If i <= UBound(thresholds) Then
            endThreshold = CDbl(thresholds(i))
        Else
            endThreshold = 999999 ' Infinity equivalent
        End If
        
        ' Determine range
        If i = 0 Then
            rangeStart = Application.WorksheetFunction.Max(start, 0)
        Else
            rangeStart = Application.WorksheetFunction.Max(start, CDbl(thresholds(i - 1)))
        End If
        rangeEnd = Application.WorksheetFunction.Min(endThreshold, start + remaining)
        hoursInRange = Application.WorksheetFunction.Max(0, rangeEnd - rangeStart)
        
        splitArray(i) = hoursInRange
        remaining = remaining - hoursInRange
        start = start + hoursInRange
        
        If remaining <= 0 Then Exit For
    Next i
    
    hoursRemain = splitArray(0)
    result("hoursRemain") = hoursRemain
    
    ' Create overtime results
    For i = 1 To UBound(splitArray)
        If i - 1 <= UBound(thresholds) And splitArray(i) > 0 Then
            key = prefix & "_OT_" & thresholds(i - 1)
            Set result(key) = CreateObject("Scripting.Dictionary")
            result(key)("wage") = CDbl(thresholdWages(i - 1))
            result(key)("hours") = splitArray(i)
            result(key)("total") = splitArray(i) * CDbl(thresholdWages(i - 1))
        End If
    Next i
    
    Debug.Print "Hours remaining for next stage: " & hoursRemain
    Set ParseOvertime = result
End Function

' Parse regular work hours
Private Function ParseRegularWork(shiftDate As Date, shiftStart As Date, remainHours As Double, breakHours As Double) As Object
    Dim result As Object
    Dim duration As Double
    Dim shiftStartDecimal As Double
    Dim shiftEndDecimal As Double
    Dim openDecimal As Double
    Dim closeDecimal As Double
    Dim timeBlocks As Collection
    Dim block As Object
    Dim allWorkSegments As Collection
    Dim segment As Object
    Dim blockDuration As Double
    Dim earlyOT As Double
    Dim lateOT As Double
    Dim workInOpening As Double
    Dim dayKey As Integer
    Dim prefix As String
    Dim remainingBreak As Double
    Dim diff As Double
    Dim i As Integer
    Dim key As String
    
    Set result = CreateObject("Scripting.Dictionary")
    Set allWorkSegments = New Collection
    
    duration = remainHours + breakHours
    shiftStartDecimal = TimeToDecimal(shiftStart)
    shiftEndDecimal = (shiftStartDecimal + duration)
    If shiftEndDecimal >= 24 Then shiftEndDecimal = shiftEndDecimal - 24
    
    openDecimal = TimeToDecimal(globalConfig.openTime)
    closeDecimal = TimeToDecimal(globalConfig.closeTime)
    If closeDecimal = 0 Then closeDecimal = 24
    
    ' Split time by day if crossing midnight
    Set timeBlocks = SplitRegularTimeByDay(shiftDate, shiftStartDecimal, shiftEndDecimal)
    
    ' Process each time block
    For i = 1 To timeBlocks.Count
        Set block = timeBlocks(i)
        blockDuration = block("endTime") - block("startTime")
        
        ' Calculate early OT, late OT, and regular work
        earlyOT = Application.WorksheetFunction.Max(openDecimal - block("startTime"), 0) - _
                  Application.WorksheetFunction.Max(openDecimal - block("endTime"), 0)
        lateOT = Application.WorksheetFunction.Max(block("endTime") - closeDecimal, 0) - _
                 Application.WorksheetFunction.Max(block("startTime") - closeDecimal, 0)
        workInOpening = blockDuration - earlyOT - lateOT
        
        ' Determine day prefix
        dayKey = Weekday(block("date"), vbMonday)
        Select Case dayKey
            Case 6: prefix = "SAT"
            Case 7: prefix = "SUN"
            Case Else: prefix = "WD"
        End Select
        
        ' Add work segments
        If workInOpening > 0 Then
            Set segment = CreateObject("Scripting.Dictionary")
            segment("name") = prefix & "_Regular"
            segment("wage") = GetWageByType(prefix & "_BASE_WAGE")
            segment("hours") = workInOpening
            allWorkSegments.Add segment
        End If
        
        If earlyOT > 0 Then
            Set segment = CreateObject("Scripting.Dictionary")
            segment("name") = prefix & "_Early_OT"
            segment("wage") = GetWageByType(prefix & "_EARLY_OT_WAGE")
            segment("hours") = earlyOT
            allWorkSegments.Add segment
        End If
        
        If lateOT > 0 Then
            Set segment = CreateObject("Scripting.Dictionary")
            segment("name") = prefix & "_Late_OT"
            segment("wage") = GetWageByType(prefix & "_LATE_OT_WAGE")
            segment("hours") = lateOT
            allWorkSegments.Add segment
        End If
    Next i
    
    ' Sort segments by wage and deduct break time from lowest wage first
    SortSegmentsByWage allWorkSegments
    remainingBreak = breakHours
    
    For i = 1 To allWorkSegments.Count
        Set segment = allWorkSegments(i)
        If remainingBreak <= 0 Then Exit For
        
        diff = segment("hours") - remainingBreak
        segment("hours") = Application.WorksheetFunction.Max(diff, 0)
        remainingBreak = -diff
    Next i
    
    ' Assemble final result
    For i = 1 To allWorkSegments.Count
        Set segment = allWorkSegments(i)
        key = segment("name")
        
        If Not result.Exists(key) Then
            Set result(key) = CreateObject("Scripting.Dictionary")
            result(key)("wage") = segment("wage")
            result(key)("hours") = 0
            result(key)("total") = 0
        End If
        
        result(key)("hours") = result(key)("hours") + segment("hours")
        result(key)("total") = result(key)("total") + segment("hours") * segment("wage")
    Next i
    
    Set ParseRegularWork = result
End Function

' Helper function to get wage by type name
Private Function GetWageByType(wageType As String) As Double
    Select Case wageType
        Case "WD_BASE_WAGE": GetWageByType = globalConfig.wdBaseWage
        Case "WD_EARLY_OT_WAGE": GetWageByType = globalConfig.wdEarlyOtWage
        Case "WD_LATE_OT_WAGE": GetWageByType = globalConfig.wdLateOtWage
        Case "SAT_BASE_WAGE": GetWageByType = globalConfig.satBaseWage
        Case "SAT_EARLY_OT_WAGE": GetWageByType = globalConfig.satEarlyOtWage
        Case "SAT_LATE_OT_WAGE": GetWageByType = globalConfig.satLateOtWage
        Case "SUN_BASE_WAGE": GetWageByType = globalConfig.sunBaseWage
        Case "SUN_EARLY_OT_WAGE": GetWageByType = globalConfig.sunEarlyOtWage
        Case "SUN_LATE_OT_WAGE": GetWageByType = globalConfig.sunLateOtWage
        Case Else: GetWageByType = 0
    End Select
End Function

' Split regular time by day for overnight shifts
Private Function SplitRegularTimeByDay(shiftDate As Date, startDecimal As Double, endDecimal As Double) As Collection
    Dim blocks As New Collection
    Dim block As Object
    Dim nextDate As Date
    
    If endDecimal <= startDecimal Then
        ' Handle overnight shift
        Set block = CreateObject("Scripting.Dictionary")
        block("date") = shiftDate
        block("startTime") = startDecimal
        block("endTime") = 24
        blocks.Add block
        
        nextDate = DateAdd("d", 1, shiftDate)
        Set block = CreateObject("Scripting.Dictionary")
        block("date") = nextDate
        block("startTime") = 0
        block("endTime") = endDecimal
        blocks.Add block
    Else
        ' Regular same-day shift
        Set block = CreateObject("Scripting.Dictionary")
        block("date") = shiftDate
        block("startTime") = startDecimal
        block("endTime") = endDecimal
        blocks.Add block
    End If
    
    Set SplitRegularTimeByDay = blocks
End Function

' Sort work segments by wage (bubble sort)
Private Sub SortSegmentsByWage(segments As Collection)
    Dim i As Integer, j As Integer
    Dim temp As Object
    Dim segment1 As Object, segment2 As Object
    
    For i = 1 To segments.Count - 1
        For j = i + 1 To segments.Count
            Set segment1 = segments(i)
            Set segment2 = segments(j)
            
            If segment1("wage") > segment2("wage") Then
                ' Swap segments
                segments.Remove i
                segments.Add segment2, , i
                segments.Remove j + 1
                segments.Add segment1, , j
            End If
        Next j
    Next i
End Sub

' =================================
' Calculation Module - VBA Version
' Converted from Google Apps Script
' =================================

Option Explicit

' Convert time to decimal hours
Public Function TimeToDecimal(timeValue As Date) As Double
    TimeToDecimal = Hour(timeValue) + Minute(timeValue) / 60
End Function

' Calculate duration between start and end times (handles overnight shifts)
Public Function GetDurationHours(startTime As Date, endTime As Date) As Double
    Dim startDecimal As Double
    Dim endDecimal As Double
    
    startDecimal = TimeToDecimal(startTime)
    endDecimal = TimeToDecimal(endTime)
    
    If endDecimal > startDecimal Then
        GetDurationHours = endDecimal - startDecimal
    Else
        ' Handle overnight shift
        GetDurationHours = endDecimal + 24 - startDecimal
    End If
End Function

' Get week range (Monday to Sunday) for a given date
Public Function GetWeekRange(inputDate As Date) As Collection
    Dim weekRange As New Collection
    Dim dayOfWeek As Integer
    Dim diffToMonday As Integer
    Dim startOfWeek As Date
    Dim endOfWeek As Date
    
    ' Get day of week (1=Sunday, 2=Monday, ..., 7=Saturday in VBA)
    dayOfWeek = Weekday(inputDate)
    
    ' Calculate difference to Monday
    If dayOfWeek = 1 Then ' Sunday
        diffToMonday = -6
    Else
        diffToMonday = 2 - dayOfWeek ' 2-2=0 for Monday, 2-3=-1 for Tuesday, etc.
    End If
    
    startOfWeek = DateAdd("d", diffToMonday, inputDate)
    endOfWeek = DateAdd("d", 6, startOfWeek)
    
    weekRange.Add startOfWeek, "startOfWeek"
    weekRange.Add endOfWeek, "endOfWeek"
    
    Set GetWeekRange = weekRange
End Function

' Round to two decimal places (more accurate than standard Round)
Public Function RoundToTwo(num As Double) As Double
    RoundToTwo = Round(num * 100 + 0.0000001) / 100
End Function

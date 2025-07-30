' =================================
' Main Processing Module - VBA Version
' Converted from Google Apps Script
' =================================

Option Explicit

' Global configuration variable
Public globalConfig As ConfigData

' Main entry point - equivalent to main() function in Apps Script
Public Sub RunPayrollCalculation()
    Dim shiftData As Collection
    Dim payslipSheet As Worksheet
    Dim inputName As String
    Dim inputDate As Date
    Dim weekRange As Collection
    Dim startOfWeek As Date
    Dim endOfWeek As Date
    Dim filteredData As Collection
    Dim groupedShiftMap As Object
    Dim summary As Object
    Dim weeklyTotal As Object
    
    ' Load global configuration
    globalConfig = GetConfig()
    If globalConfig.openTime = 0 Then
        MsgBox "Failed to load configuration. Cannot proceed."
        Exit Sub
    End If
    
    ' Get validated shift data
    Set shiftData = GetValidatedShiftData()
    If shiftData.Count = 0 Then
        MsgBox "No valid shift data found."
        Exit Sub
    End If
    
    ' Get input parameters from Payslip sheet
    On Error GoTo ErrorHandler
    Set payslipSheet = ThisWorkbook.Worksheets("Payslip")
    inputName = Trim(CStr(payslipSheet.Range("A1").Value))
    inputDate = CDate(payslipSheet.Range("B1").Value)
    
    If inputName = "" Then
        MsgBox "Please enter employee name in cell A1 of Payslip sheet."
        Exit Sub
    End If
    
    ' Get week range
    Set weekRange = GetWeekRange(inputDate)
    startOfWeek = weekRange("startOfWeek")
    endOfWeek = weekRange("endOfWeek")
    
    ' Filter data for the specified employee and week
    Set filteredData = FilterShiftData(shiftData, inputName, startOfWeek, endOfWeek)
    
    If filteredData.Count = 0 Then
        MsgBox "No shift data found for " & inputName & " during the week of " & FormatDate(startOfWeek) & " to " & FormatDate(endOfWeek)
        Exit Sub
    End If
    
    ' Sort and group data by date
    Set groupedShiftMap = SortAndGroupByDate(filteredData)
    
    ' Parse shifts and calculate wages
    Set summary = ParseShift(groupedShiftMap)
    
    ' Calculate weekly totals
    Set weeklyTotal = CalculateWeeklyTotal(summary)
    
    ' Debug output
    Debug.Print "=== Payroll Calculation Results ==="
    Debug.Print "Employee: " & inputName
    Debug.Print "Week: " & FormatDate(startOfWeek) & " to " & FormatDate(endOfWeek)
    Debug.Print "Total Hours: " & weeklyTotal("hours")
    Debug.Print "Total Pay: $" & RoundToTwo(weeklyTotal("total"))
    Debug.Print "================================="
    
    ' Render payslip
    RenderPaySlip inputName, startOfWeek, endOfWeek, groupedShiftMap, summary, weeklyTotal
    
    MsgBox "Payroll calculation completed successfully for " & inputName & vbCrLf & _
           "Total Hours: " & RoundToTwo(weeklyTotal("hours")) & vbCrLf & _
           "Total Pay: $" & RoundToTwo(weeklyTotal("total")), vbInformation, "Payroll Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in payroll calculation: " & Err.Description, vbCritical, "Error"
End Sub

' Filter shift data for specific employee and date range
Private Function FilterShiftData(shiftData As Collection, employeeName As String, startDate As Date, endDate As Date) As Collection
    Dim filteredData As New Collection
    Dim i As Integer
    Dim record As ShiftRecord
    
    For i = 1 To shiftData.Count
        record = shiftData(i)
        If record.employeeName = employeeName And _
           record.shiftDate >= startDate And _
           record.shiftDate <= endDate Then
            filteredData.Add record
        End If
    Next i
    
    Set FilterShiftData = filteredData
End Function

' Calculate weekly total from summary
Private Function CalculateWeeklyTotal(summary As Object) As Object
    Dim weeklyTotal As Object
    Dim key As Variant
    Dim hours As Double
    Dim total As Double
    
    Set weeklyTotal = CreateObject("Scripting.Dictionary")
    hours = 0
    total = 0
    
    For Each key In summary.Keys
        hours = hours + summary(key)("hours")
        total = total + RoundToTwo(summary(key)("wage") * summary(key)("hours"))
    Next key
    
    weeklyTotal("hours") = RoundToTwo(hours)
    weeklyTotal("total") = RoundToTwo(total)
    
    Set CalculateWeeklyTotal = weeklyTotal
End Function

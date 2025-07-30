' =================================
' Utility Functions Module - VBA Version
' Converted from Google Apps Script
' =================================

Option Explicit

' Display errors to user
Public Sub RaiseErrors(message As String, errors As Collection)
    Dim errorMsg As String
    Dim i As Integer
    
    errorMsg = message
    For i = 1 To errors.Count
        errorMsg = errorMsg & errors(i) & vbCrLf
    Next i
    
    ' Log to Immediate window (equivalent to Logger.log)
    Debug.Print errorMsg
    
    ' Show message box to user
    MsgBox errorMsg, vbExclamation, "Payroll System Errors"
End Sub

' Debug function to print collections/dictionaries (equivalent to logOutMap)
Public Sub LogOutCollection(inputCollection As Collection, Optional keyName As String = "")
    Dim item As Variant
    Dim i As Integer
    
    Debug.Print "=== Collection: " & keyName & " ==="
    
    For i = 1 To inputCollection.Count
        Set item = inputCollection(i)
        If TypeName(item) = "Dictionary" Then
            LogOutDictionary item
        Else
            Debug.Print "Item " & i & ": " & item
        End If
    Next i
    
    Debug.Print "=== End Collection ==="
End Sub

' Debug function to print dictionary contents
Public Sub LogOutDictionary(inputDict As Object)
    Dim key As Variant
    
    Debug.Print "--- Dictionary Contents ---"
    For Each key In inputDict.Keys
        Debug.Print key & ": " & inputDict(key)
    Next key
    Debug.Print "--- End Dictionary ---"
End Sub

' Format date for display (equivalent to Utilities.formatDate)
Public Function FormatDate(dateValue As Date) As String
    FormatDate = Format(dateValue, "dd/mm/yyyy")
End Function

' Format time for display
Public Function FormatTime(timeValue As Date) As String
    FormatTime = Format(timeValue, "h:mm AM/PM")
End Function

' Get day name from date
Public Function GetDayName(dateValue As Date) As String
    Dim dayNames As Variant
    dayNames = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    GetDayName = dayNames(Weekday(dateValue) - 1)
End Function

' Convert array to VBA array (helper function)
Public Function ArrayToVBA(inputArray As Variant) As Variant
    Dim i As Integer
    Dim result() As Double
    
    ReDim result(0 To UBound(inputArray))
    
    For i = 0 To UBound(inputArray)
        result(i) = CDbl(inputArray(i))
    Next i
    
    ArrayToVBA = result
End Function

' Check if a value exists in an array
Public Function InArray(valueToFind As Variant, arrayToSearch As Variant) As Boolean
    Dim i As Integer
    InArray = False
    
    For i = LBound(arrayToSearch) To UBound(arrayToSearch)
        If arrayToSearch(i) = valueToFind Then
            InArray = True
            Exit Function
        End If
    Next i
End Function

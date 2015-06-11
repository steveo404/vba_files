Attribute VB_Name = "modTodayDate"
Option Compare Database

Function TodayDate()
    'Function pulls the current day's date
    'Returns the date as a String in MMDDYY format

    Dim today As String
    Dim mnth As String
    Dim yr As String
    
    today = Day(Date)
    mnth = Month(Date)
    yr = Year(Date)
    
    If Len(mnth) = 1 Then
        mnth = "0" & mnth
    End If
    If Len(today) = 1 Then
        today = "0" & today
    End If
    
    TodayDate = mnth + today + yr

End Function

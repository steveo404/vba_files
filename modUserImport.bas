Attribute VB_Name = "modUserImport"
Option Compare Database


Sub UserImport()
    'Script used to import files from net user command
    'Command to generate the text file is <net user /domain > c:\users.txt>

    Dim strFile As String
    Dim entry As String
    Dim check As String
    Dim trigger As Integer
    Dim strLength As Integer
    Dim i, j As Integer
    Dim userID As String
    Dim resultFile As String
    Dim scriptFile As String
    
    
    'strFile = GetFileName()

    'Remove the resultFile from a previous run of this script
    resultFile = "C:\UserResults.txt"
    Open resultFile For Append As #9
    Close #9
    Kill resultFile
    
    'Remove the resultFile from a previous run of this script
    scriptFile = "C:\Scripts\UsersForScript.txt"
    Open scriptFile For Append As #99
    Close #99
    Kill scriptFile
    
    
    trigger = 0
    
    Open strFile For Input As #1    'Open the text file
    Do Until EOF(1)
        Input #1, entry
        check = Mid(entry, 1, 5)

        If trigger > 0 Then        'Once you have a device name you can look for the IP Address - this can only be done after a device name is populated
            If entry = "The command completed successfully." Then Exit Do
            strLength = Len(Trim(entry))
            If strLength > 57 Then
                j = 1
                For i = 1 To 3
                    userID = Trim(Mid(entry, j, 25))
                    Open resultFile For Append As #9
                    Write #9, userID
                    Close #9
                    Open scriptFile For Append As #99
                    Print #99, userID
                    Close #99
                    j = j + 25
                Next
            End If
            If (strLength > 31) And (strLength < 51) Then
                j = 1
                For i = 1 To 2
                    userID = Trim(Mid(entry, j, 25))
                    Open resultFile For Append As #9
                    Write #9, userID
                    Close #9
                    Open scriptFile For Append As #99
                    Print #99, userID
                    Close #99
                    j = j + 25
                Next
            End If
            If strLength < 26 Then
                    userID = Trim(Mid(entry, j, 25))
                    Open resultFile For Append As #9
                    Write #9, userID
                    Close #9
                    Open scriptFile For Append As #99
                    Print #99, userID
                    Close #99
            End If
        End If
        
        If check = "-----" Then        'Review text in the file to find the Name field which represents name of the device
            trigger = 1
        End If
    Loop
        
    
    Dim strTableName As String
    Dim fileDate As String
    
    fileDate = TodayDate()
    
    'Remove previous version of this file if it was run earlier in the dat
    'It now has a date stamp included with the file name
    
    strTableName = "UserIDLookupResults" + fileDate
    If TableExists(strTableName) Then
        DoCmd.DeleteObject acTable, strTableName
    End If
    
    'Import the text file into a database table
    DoCmd.TransferText acImportDelim, , strTableName, resultFile, False
    
    Application.RefreshDatabaseWindow
    
    MsgBox "Script Complete!!!"
    

End Sub

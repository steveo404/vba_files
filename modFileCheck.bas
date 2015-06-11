Attribute VB_Name = "modFileCheck"
Option Compare Database

Sub FileCheck()

    Dim strPath As String
    Dim strFile As String
    Dim i As Integer
    
    'strPath = "C:\Trigger\"
    'strFile = Dir(strPath & "*.*")
    
    i = 0
    
    Do
        Open "C:\Trigger\File" & i & ".txt" For Append As #3
        Close #3
        i = i + 1
        If i = 25 Then
            Open "C:\Trigger\helloworld.txt" For Append As #4
            Close #4
        End If
        If Len(Dir("c:\Trigger\helloworld.txt")) <> 0 Then
            Exit Do
        End If

    Loop
    
    



End Sub


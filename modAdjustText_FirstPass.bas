Attribute VB_Name = "modAdjustText_FirstPass"
Option Compare Database
Option Explicit
Sub AdjustmentTextReformat()

    Dim fileName As String
    Dim outFileName As String
    Dim entry As String
    Dim checkEntry As String
    Dim holdEntry As String
    Dim grabEntry As String
    Dim headerCheck As Integer
    Dim fileSet As Integer
    
    outFileName = "C:\Users\soneal\Documents\Data\TEST_output220_FirstPass.txt"
    
    Open outFileName For Append As #1
    Close #1
    Kill outFileName
    
    fileName = "C:\Users\soneal\Documents\Data\ARAdjustments.txt"
    
    headerCheck = 1

    Open fileName For Input As #3
    Do Until EOF(3)
        Input #3, entry
        grabEntry = Left(entry, 12)
        If grabEntry <> "HNTB Invoice" Then
            If Left(grabEntry, 4) <> "Page" And Left(grabEntry, 4) <> "http" Then
                If Len(grabEntry) > 0 Then
                    Open outFileName For Append As #44
                    Print #44, entry
                    Close #44
                End If
            End If
        End If
    Loop

End Sub

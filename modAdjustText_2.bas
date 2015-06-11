Attribute VB_Name = "modAdjustText_2"
Option Compare Database
Option Explicit
Sub AdjustmentTextReformat()

    Dim fileName As String
    Dim outFileName As String
    Dim entry As String
    Dim checkEntry As String
    Dim holdEntry As String
    Dim grabEntry As String
    Dim skipEntry As String
    Dim headerCheck As Integer
    Dim fileSet As Integer
    
    outFileName = "C:\Users\soneal\Documents\Data\TEST_output.txt"
    
    Open outFileName For Append As #1
    Close #1
    Kill outFileName
    
    fileName = "C:\Users\soneal\Documents\Data\07272014 to 12262014 Adjustment log.txt"
    
    headerCheck = 1

    Open fileName For Input As #3
    Do Until EOF(3)
        Input #3, entry
        grabEntry = Replace(entry, vbTab, "") 'Grab the text, strip tabs and assign it to the grabEntry variable
        skipEntry = Mid(grabEntry, 1, 4)
        If skipEntry <> "Page" And skipEntry <> "HNTB" And skipEntry <> "" Then
            checkEntry = Mid(entry, 1, 2)
            If checkEntry = "01" Or checkEntry = "12" Then
                Open outFileName For Append As #99
                Print #99, holdEntry
                Close #99
                holdEntry = grabEntry
            End If
            If checkEntry <> "01" Then
                holdEntry = holdEntry + " " + grabEntry
            End If
        End If
    Loop

End Sub

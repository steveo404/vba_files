Attribute VB_Name = "modAdjustText_SecondPass"
Option Compare Database
Option Explicit
Sub AdjustmentTextReformat_2()

    Dim fileName As String
    Dim outFileName As String
    Dim entry As String
    Dim checkEntry As String
    Dim holdEntry As String
    Dim grabEntry As String
    Dim headerCheck As Integer
    Dim fileSet As Integer
    
    Reset
    
    outFileName = "C:\Users\soneal\Documents\Data\TEST_output220_SecondPass.txt"
    
    Open outFileName For Append As #1
    Close #1
    Kill outFileName
    
    fileName = "C:\Users\soneal\Documents\Data\TEST_output220_FirstPass.txt"
    
    headerCheck = 1

    Open fileName For Input As #3
    Do Until EOF(3)
        Input #3, entry
        grabEntry = Replace(entry, vbTab, "") 'Grab the text, strip tabs and assign it to the grabEntry variable
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
    Loop
    Open outFileName For Append As #44
    Print #44, holdEntry
    Close #44
End Sub

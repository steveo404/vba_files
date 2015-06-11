Attribute VB_Name = "modAdjustText_ThirdPass"
Option Compare Database
Option Explicit
Sub AdjustmentTextReformat_3()

    Dim fileName As String
    Dim outFileName As String
    Dim entry As String
    Dim checkEntry As String
    Dim linePlace As String
    Dim grabEntry As String
    Dim restLine As String
    Dim firm As String
    Dim div As String
    Dim BO As String
    Dim amount As String
    Dim curr As String
    Dim reason As String
    Dim glDate As String
    Dim transNumber As String
    Dim job As String
    Dim comment As String
    
    Dim dummy As Integer
    
        
    Reset
    
    outFileName = "C:\Users\soneal\Documents\Data\TEST_output220_ThirdPass.txt"
    
    Open outFileName For Append As #1
    Close #1
    Kill outFileName
    
    fileName = "C:\Users\soneal\Documents\Data\TEST_output220_SecondPass.txt"
    
    Open fileName For Input As #3
    Do Until EOF(3)
        Input #3, entry
        checkEntry = Mid(entry, 1, 2)
        If checkEntry = "01" Or checkEntry = "12" Then
            firm = Mid(entry, 1, 2)
            div = Mid(entry, 4, 2)
            BO = Mid(entry, 7, 3)
            restLine = Trim(Mid(entry, 10, Len(entry) - 9))
            amount = Mid(restLine, 1, InStr(restLine, " "))
            restLine = Trim(Mid(restLine, InStr(restLine, " "), Len(restLine) - Len(amount) + 1))
            curr = Mid(restLine, 1, 3)
            restLine = Trim(Mid(restLine, 4, Len(restLine) - Len(curr)))
            linePlace = GetPositionOfFirstNumericCharacter(restLine)
            reason = Mid(restLine, 1, linePlace - 1)
            restLine = Trim(Mid(restLine, linePlace, Len(restLine) - Len(reason)))
            glDate = Mid(restLine, 1, 9)
            restLine = Trim(Mid(restLine, 10, Len(restLine) - Len(glDate)))
            transNumber = Mid(restLine, 1, InStr(restLine, " "))
            restLine = Trim(Mid(restLine, InStr(restLine, " "), Len(restLine) - Len(transNumber) + 1))
            job = Mid(restLine, 1, 5)
            dummy = Len(restLine)
            comment = Trim(Mid(restLine, 7, Len(restLine) - 6))
            Open outFileName For Append As #99
            Write #99, firm, div, BO, amount, curr, reason, glDate, transNumber, job, comment
            Close #99
        End If
    Loop
    
    MsgBox ("Third Pass Complete")

End Sub


Public Function GetPositionOfFirstNumericCharacter(ByVal s As String) As Integer
    Dim i As Integer
    
    For i = 1 To Len(s)
        Dim currentCharacter As String
        currentCharacter = Mid(s, i, 1)
        If IsNumeric(currentCharacter) = True Then
            GetPositionOfFirstNumericCharacter = i
            Exit Function
        End If
    Next i
End Function

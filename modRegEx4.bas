Attribute VB_Name = "modRegEx4"
Option Compare Database

Private Sub userIDRegex3()
    Dim strPattern As String: strPattern = "[^a-zA-Z]"
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    
    strInput = "Regiona_lConnect"

    If strPattern <> "" Then
        strReplace = ""

        With regEx
            .Global = True
            .Multiline = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
            'MsgBox (regEx.Replace(strInput, strReplace))
            MsgBox ("Matched")
        Else
            MsgBox ("Not matched")
        End If
    End If
End Sub



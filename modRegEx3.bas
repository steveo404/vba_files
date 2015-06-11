Attribute VB_Name = "modRegEx3"
Option Compare Database

Private Sub emailRegex2()
    Dim strPattern As String: strPattern = "[a-z,A-Z]*@[a-z,A-Z]*.com"
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    
    strInput = "ABC@abc.com"

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


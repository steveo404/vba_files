Attribute VB_Name = "modRegEx1"
Option Compare Database

Sub RegEx_Tester()
    'Dim strPattern As String: strPattern = "[a-z,A-Z]"
    
    Set objRegExp_1 = CreateObject("vbscript.regexp")
    objRegExp_1.Global = True
    objRegExp_1.IgnoreCase = True
    'objRegExp_1.Pattern = strPattern
    'objRegExp_1.Pattern = [a-z,A-Z]
    strToSearch = "ABC@xyz.com"
    
    Set regExp_Matches = objRegExp_1.Execute(strToSearch)
    
    If regExp_Matches.Count = 1 Then
    MsgBox ("This string is a valid email address.")
    End If
End Sub


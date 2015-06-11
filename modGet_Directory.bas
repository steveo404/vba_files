Attribute VB_Name = "modGet_Directory"
Option Compare Database

Function Get_Directory(ByRef strMessage As String) As String
    'Function allows users to select a directory
    'This function is specifically designed to select directory on I:\ drive
    'Returns the directory path as a string
    
    On Error GoTo BadDirections
    
    Dim objFolderRef As Object
    Set objFolderRef = CreateObject("Shell.Application").BrowseForFolder _
    (0, strMessage, &H4000, "C:\")
    If Not objFolderRef Is Nothing Then
        Get_Directory = objFolderRef.items.Item.path
    Else
        Get_Directory = vbNullString
    End If
    
    Set objFoderRef = Nothing
    Exit Function
    
BadDirections:
    Set objFolderRef = Nothing
    Get_Directory = "Error Selecting a Folder"

End Function

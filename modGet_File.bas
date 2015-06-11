Attribute VB_Name = "modGet_File"
Option Compare Database

Public Function GetFileName()
    'Function allows users to select a directory and a file
    'Requires following references
    'Visual Basic for Applications
    'Microsoft Access 12.0 Object Library
    'OLE Automation
    'Microsoft Visual Basic for Applications Extensibility
    'Microsoft ActiveX Data Objects 2.1 Library
    'Microsoft Office 12 Object Library
    
    Dim result As Integer
    Dim fileName As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select file"
        .Filters.Add "All Files", "*.*"
        .Filters.Add "Excel Files", "*.xlsx"
        .AllowMultiSelect = False
        .InitialFileName = CurrentProject.path
        
        result = .Show
        If (result <> 0) Then
            fileName = Trim(.SelectedItems.Item(1))
        End If
    End With
    
    GetFileName = fileName
    
End Function

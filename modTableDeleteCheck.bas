Attribute VB_Name = "modTableDeleteCheck"
Option Compare Database

Public Function TableExists(sTable As String) As Boolean
    Dim tdf As TableDef
    
    On Error Resume Next
    
    Set tdf = CurrentDb.TableDefs(sTable)
    
    If Err.Number = 0 Then
        TableExists = True
    Else
        TableExists = False
    End If
    
End Function

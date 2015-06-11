Attribute VB_Name = "modExportDBObjects"
Option Compare Database
Option Explicit

Public Sub ExportModCode()
'Script Name:       ExportModCode
'Author:            Steve O'Neal
'Created:           11/06/2014
'Last Modified:     2/2/2015
'Version:           1.0
'Dependency:        NONE
'
'Script used to export modules, scripts, and queries as text files
'Files are exported to 'Source' folder on C:\ drive by database name


On Error GoTo Err_ExportModCode
    
    Dim db As Database
    'Dim db As DAO.Database
    Dim td As TableDef
    Dim dbName As String
    Dim d As Document
    Dim c As Container
    Dim i As Integer
    Dim sExportLocation As String
    
    Set db = CurrentDb()
    
    dbName = Application.CurrentProject.Name
    
    sExportLocation = "C:\Source\" & dbName & "\" 'Do not forget the closing back slash! ie: C:\Temp\

    'FolderCreate sExportLocation
    
    Set c = db.Containers("Forms")
    For Each d In c.Documents
        Application.SaveAsText acForm, d.Name, sExportLocation & "Form_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Reports")
    For Each d In c.Documents
        Application.SaveAsText acReport, d.Name, sExportLocation & "Report_" & d.Name & ".txt"
    Next d
   
    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        Application.SaveAsText acMacro, d.Name, sExportLocation & "Macro_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Modules")
    For Each d In c.Documents
        Application.SaveAsText acModule, d.Name, sExportLocation & "Module_" & d.Name & ".txt"
    Next d
    
    For i = 0 To db.QueryDefs.Count - 1
        Application.SaveAsText acQuery, db.QueryDefs(i).Name, sExportLocation & "Query_" & db.QueryDefs(i).Name & ".txt"
    Next i
    
    Set db = Nothing
    Set c = Nothing
    
    MsgBox "All database objects have been exported as a text file to " & sExportLocation, vbInformation
    
Exit_ExportModCode:
    Exit Sub
    
Err_ExportModCode:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ExportModCode
    
End Sub

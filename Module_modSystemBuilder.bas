Private Const strModuleName As String = "modSystemBuilder"
Option Compare Database
Option Explicit

'===
'This routine exports all database objects to a text file
'===
Public Sub ExportDatabaseObjects()
On Error GoTo Err_ExportDatabaseObjects
    
    Dim db As Database
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim i As Integer
    Dim sExportLocation As String
    
    Set db = CurrentDb()
    
    sExportLocation = "C:\Users\Joe Feldmann\Documents\Development\Portfolio_Management_Tool\Test\"
    
    For Each td In db.TableDefs 'Tables
        If Left(td.Name, 4) <> "MSys" Then
            DoCmd.TransferText acExportDelim, , td.Name, sExportLocation & "Table_" & td.Name & ".txt", True
        End If
    Next td
    
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
        Application.SaveAsText acModule, d.Name, sExportLocation & "Module_" & d.Name & ".bas"
    Next d
    
    For i = 0 To db.QueryDefs.Count - 1
        Application.SaveAsText acQuery, db.QueryDefs(i).Name, sExportLocation & "Query_" & db.QueryDefs(i).Name & ".txt"
    Next i
    
    Set db = Nothing
    Set c = Nothing
    
    MsgBox "All database objects have been exported as a text file to " & sExportLocation, vbInformation
    
Exit_ExportDatabaseObjects:
    Exit Sub
    
Err_ExportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ExportDatabaseObjects
    
End Sub

'===
'This routine imports any database object stored in a text file. The naming convention of the text file must match what the routine is expecting.
'It will not overwrite the module where this routine resides.
'===
Public Sub ImportDatabaseObjects()
    Dim objFSO As New Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim strImportPath As String
    Dim strFileName As String

    ''' NOTE: Path where the code modules are located.
    strImportPath = "C:\Users\Joe Feldmann\Documents\Development\Portfolio_Management_Tool\Test\"
        
    If objFSO.GetFolder(strImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    'Call DeleteVBAModulesAndUserForms
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(strImportPath).Files
        Select Case Left(objFile.Name, InStr(1, objFile.Name, "_") - 1)
            Case "Form"
                Application.LoadFromText acForm, Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1, Len(Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1)) - 4), objFile.Path
            Case "Module"
                If Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1, Len(Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1)) - 4) <> strModuleName Then
                    Application.LoadFromText acModule, Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1, Len(Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1)) - 4), objFile.Path
                End If
            Case "Query"
                Application.LoadFromText acQuery, Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1, Len(Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1)) - 4), objFile.Path
            Case "Table"
                Application.LoadFromText acTable, Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1, Len(Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1)) - 4), objFile.Path
            Case "Report"
                Application.LoadFromText acReport, Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1, Len(Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1)) - 4), objFile.Path
            Case "Macro"
                Application.LoadFromText acMacro, Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1, Len(Mid(objFile.Name, InStr(1, objFile.Name, "_") + 1)) - 4), objFile.Path
            
        End Select
    Next objFile
    
    MsgBox "Import is ready"
End Sub
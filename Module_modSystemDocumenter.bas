Private Const strModuleName As String = "modSystemDocumenter"
Option Compare Database

'===
'This routine counts the number of lines in a VBA routine
'===
Function countSubLines(CodeMod As Object, lineNumber As Long)
    Dim i As Integer
    Dim currentLine As String
    
    'Set the starting point
    currentLine = CodeMod.Lines(lineNumber + i, 1)
    
    i = 1
    
    'Loop until the final line of the routine is reached
    Do Until InStr(1, currentLine, "End Sub") Or _
            InStr(1, currentLine, "End Function") Or _
            InStr(1, currentLine, "End Property") Or _
            InStr(1, currentLine, "End Type")
        
        currentLine = CodeMod.Lines(lineNumber + i, 1)
        i = i + 1
        
    Loop
    
    'Return the number of lines that were iterated over
    countSubLines = i
      
End Function
 
'===
'This routine finds the comments above a routine and parses out the description text
'===
Function RE6(strData As String) As String
    Dim CanDo As Boolean
    Dim RE
    
    Set RE = New RegExp
    
    With RE
        .MultiLine = True
        .Global = True
        .IgnoreCase = True
        .Pattern = "===$([\s\S]*?)==="
    End With
    
    If RE.Test(strData) = True Then CanDo = True
    
    If CanDo Then
        Set REMatches = RE.Execute(strData)
        RE6 = REMatches.Item(0).SubMatches.Item(0)
    Else
        RE6 = ""
    End If
 
End Function

'===
'This procedure writes all the information about the VBA code into an Excel workbook
'===
Sub InsertProcedureNameIntoProcedures()
    Const conProcName As String = "InsertProcedureNameIntoProcedures"
    Dim ProcName As String
    Dim ProcLine As String
    Dim StartLine As Long
    Dim ProcType As VBIDE.vbext_ProcKind
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim Done As Boolean
    Dim ProcBodyLine As Long
    Dim SaveProcName As String
    Dim ConstName As String
    Dim ValidConstName As Boolean
    Dim ConstAtLine As Long
    Dim EndOfDeclaration As Long
    Dim strDeclaration As String
    Dim i As Integer
    Dim objExcel As Excel.Application
    Dim wbk As Workbook
    Dim wks As Worksheet
    Dim CodeContent As String
    Dim CodeDescription As String
    Dim NegativeOffset As Integer
            
    'Open a workbook to capture the output
    Set objExcel = fnGetExcel
    Set wbk = objExcel.Workbooks.Add
    Set wks = wbk.ActiveSheet
    objExcel.Visible = True
  
    'Iterate through all the modules within the VB Project
    For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
        'Skip the soft deleted modules
        If Left(VBComp.Name, 6) = "Form_z" Then GoTo Skip_vbComp
        
        'Set initial objects
        Set CodeMod = VBComp.CodeModule
        StartLine = CodeMod.CountOfDeclarationLines + 1
        NegativeOffset = StartLine
        
        'Iterate through each line of the code
        Do Until Done
            'Initialize string variables
            CodeContent = ""
            CodeDescription = ""
            
            'Determine the routine name
            ProcName = CodeMod.ProcOfLine(StartLine, ProcType)
            If Len(ProcName) = 0 Then GoTo Skip_vbComp
            
            'Get the code for the routine
            ProcBodyLine = CodeMod.ProcBodyLine(ProcName, ProcType)
            CodeContent = CodeMod.Lines(ProcBodyLine, countSubLines(CodeMod, ProcBodyLine))
            
            'Get the descriptions
            If ProcBodyLine - NegativeOffset > 0 Then CodeDescription = CodeMod.Lines(ProcBodyLine - NegativeOffset, NegativeOffset)
            If Len(CodeDescription) > 0 Then CodeDescription = RE6(CodeDescription)
            StartLine = ProcBodyLine + CodeMod.ProcCountLines(ProcName, ProcType) + 1
            NegativeOffset = CodeMod.ProcCountLines(ProcName, ProcType)
            
            'Check to see if routine has been documented already
            If ProcName = SaveProcName Then
                Done = True
            Else
                'Write information to Excel file
                SaveProcName = ProcName
                i = i + 1
                If Left(SaveProcName, 9) = "subSelect" Then Stop
                wks.Cells(i, 1).Value = VBComp.Name
                wks.Cells(i, 2).Value = SaveProcName
                wks.Cells(i, 3).Value = CodeDescription
                wks.Cells(i, 4).Value = CodeContent
            End If
        Loop
        
        'Clean up
        Set CodeMod = Nothing
        Done = False
        
Skip_vbComp:
    Next VBComp
    Set VBComp = Nothing
    Set CodeMod = Nothing
    Set wks = Nothing
    Set wbk = Nothing
    Set objExcel = Nothing
End Sub
Private Const strModuleName As String = "modErrorHandler"
Option Compare Database

'===
'Error handler
'===

Function fncErrorHandler(ByVal strRoutineName As String, ByVal strModuleName As String, ByVal lngLineNumber As Long, _
                        ByVal blnDisplay As Boolean, Optional ByVal strErrorDescription As String) As Boolean
    On Error GoTo Err_fncErrorHandler
    Dim lngErrorID As Long
    Const strErrorProcedure As String = "sp_ErrorInsert"
    Const conRoutineName As String = "fncErrorHandler"
    
    'Pass parameters to stored procedure to log the error centrally
    'lngErrorID = fncInsertRecord(strErrorProcedure, strModuleName, strRoutineName, lngLineNumber, strErrorDescription)
    
    'Check results of error logging
    If blnDisplay Then
        If lngErrorID = 0 Then
            MsgBox "The application has encountered an error and cannot complete the requested action." & _
                    vbCrLf & vbCrLf & _
                    "The application was not able to log the error.", vbOKOnly + vbCritical
        ElseIf lngErrorID > 0 Then
            MsgBox "The application has encountered an error and cannot complete the requested action." & _
                    vbCrLf & vbCrLf & _
                    "The Error # is: " & lngErrorID & ". Please contact system support and provide this Error #.", vbOKOnly + vbCritical
        End If
    End If
    
    fncErrorHandler = True
    
Exit_fncErrorHandler:
    Exit Function
Err_fncErrorHandler:
    fncErrorHandler = False
    MsgBox "The Error Handler has encountered an error." & _
    vbCrLf & _
    vbCrLf & _
    "Please contact System Support!", vbOKOnly + vbExclamation
    
    GoTo Exit_fncErrorHandler
    
End Function
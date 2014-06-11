Attribute VB_Name = "modPopulateFormData"
Private Const strModuleName As String = "modPopulateFormData"
Option Compare Database
Option Explicit

'===
'Populates the values for all controls on a form that align with the passed recordset
'===
Sub subPopulateControls(ByVal frmTarget As Form, ByVal rstSource As ADODB.Recordset)
    On Error GoTo Err_Sub
    Dim ctlValue As Access.Control
    Const conProcName = "subPopulateControls"
    
    If rstSource.EOF Then
        MsgBox "No data for selected record.", vbExclamation, "No Such Record"
        GoTo Exit_Sub
    Else
        On Error Resume Next
        'Loop through controls on the form
        For Each ctlValue In frmTarget.Controls
            'Perform different actions for different types of controls
            Select Case ctlValue.ControlType
                Case acLabel, acCommandButton, acSubform:
                Case Else
                    'Set the Value property of the control
                    ctlValue.Value = rstSource.Fields(ctlValue.Tag).Value
            End Select
        Next ctlValue
        
        On Error GoTo Err_Sub
    End If
    
Exit_Sub:
    Set ctlValue = Nothing
    Exit Sub
Err_Sub:
    If fncErrorHandler(conProcName, strModuleName, Erl, True, Err.Description) Then
        GoTo Exit_Sub
    Else
        Resume Next
    End If
End Sub

'===
'This routine determines all the lists that must be populated and gets the values and writes them to the controls
'===
Sub subPopulateFormLists(ByVal frmLoad As Access.Form)
    On Error GoTo Err_Sub
    Dim rstLists As ADODB.Recordset
    Dim rstSingleList As ADODB.Recordset
    Dim ctlList As Access.Control
    Const strFormListsProcedure As String = "qFormListsSelect"
    Const conProcName As String = "subPopulateFormLists"
    
    'Determine all the lists that need to be populated
    'Set rstLists = fncselectdataset(strFormListsProcedure, True, frmload.name)
    
    Do Until rstLists.EOF
        'Get the list of values
        Set rstSingleList = fncselectdataset(rstLists.Fields(1).Value, False)
        'Get the control where the list will be populated
        Set ctlList = frmLoad.Controls(rstLists.Fields(2).Value)
        'Put the recordset in the list
        Call subPopulateControlWithRecordset(rstSingleList, ctlList)
        
        rstLists.MoveNext
    Loop
    
Exit_Sub:
    Set ctlList = Nothing
    Set rstSingleList = Nothing
    Set rstLists = Nothing
    Exit Sub
Err_Sub:
    If fncErrorHandler(conProcName, strModuleName, Erl, True, Err.Description) Then
        GoTo Exit_Sub
    Else
        Resume Next
    End If
End Sub

'===
'This routine clears all selections from a list box
'===
Sub subUnselectListBoxItems(ctlList As ListBox)
    On Error GoTo Err_Sub
    Dim varItem As Variant
    Const conProcName As String = "subUnselectListBoxItems"
    
    For Each varItem In ctlList.ItemsSelected
        ctlList.Selected(varItem) = False
    Next
    
Exit_Sub:
    Exit Sub
Err_Sub:
    If fncErrorHandler(conProcName, strModuleName, Erl, True, Err.Description) Then
        GoTo Exit_Sub
    Else
        Resume Next
    End If
End Sub

'===
'This routine deletes the list of values in a control
'===
Sub subDeleteControlValueList(ctlList As Access.Control, ByVal blnDeleteAll As Boolean)
    On Error GoTo Err_Sub
    Dim iList As Integer
    Const conProcName As String = "subDeleteControlValueList"
    
    If blnDeleteAll Then
        For iList = 0 To ctlList.ListCount - 1
            ctlList.RemoveItem 0
        Next iList
    Else
        ctlList.RemoveItem ctlList.ItemsSelected(0)
    End If
    
Exit_Sub:
    Exit Sub
Err_Sub:
    If fncErrorHandler(conProcName, strModuleName, Erl, True, Err.Description) Then
        GoTo Exit_Sub
    Else
        Resume Next
    End If
End Sub

'===
'This routine populates a control with the values in a recordset
'===
Sub subPopulateControlWithRecordset(ByVal rstSource As ADODB.Recordset, ByVal ctlTarget As Access.Control)
On Error GoTo Err_Sub
    Dim fld As ADODB.Field
    Const conProcName As String = "subPopulateControlWithRecordset"
    
    'Clear the list in advance
    Call subDeleteControlValueList(ctlTarget, True)
    
    With rstSource
        If .RecordCount > 0 Then
        .MoveFirst
            'If there is data, then loop through every row of the recordset
            Do Until .EOF
                strData = ""
                'Concatenate a string of the data from each field
                For Each fld In .Fields
                    If strData = "" Then
                        strData = Chr(39) & fld.Value & Chr(39)
                    Else
                        strData = strData & ";" & Chr(39) & fld.Value & Chr(39)
                    End If
                Next fld
                
                'Add the value to the list
                ctlTarget.AddItem strData
                
                .MoveNext
            Loop
        End If
    End With
    
Exit_Sub:
    Set fld = Nothing
    Exit Sub
Err_Sub:
    If fncErrorHandler(conProcName, strModuleName, Erl, True, Err.Description) Then
        GoTo Exit_Sub
    Else
        Resume Next
    End If
End Sub

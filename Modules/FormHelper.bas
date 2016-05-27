Attribute VB_Name = "FormHelper"
' Module        : FormHelper
' Description   : Helper module to store shared procedures
' Author        : Azri Ahmad Rosehaizat
' Created       : May 2016
' --------------------------------------------------
Option Compare Database
Option Explicit

'Switch between View/Edit mode
Public Sub ControlSwitch(flag As Boolean, tag As String, currFrm As Form)
If gcvHandleError Then On Error GoTo PROC_EXIT
'    Purpose:   Dynamically change the UI to edit/read mode
'    Params:    flag - decision maker
'               tag - The property of .tag field of a control
'               currFrm - The form to check
    
    Dim frm As Form
    Dim ctl As Control
    
    
    Set frm = currFrm           'Set the form

    'Loop through all the controls in the form and set value of certain properties
    For Each ctl In frm.Controls
        If ctl.tag = tag Then
            Select Case ctl.ControlType
                Case acTextBox, acComboBox
                    'Dynamic switcher
                    ctl.BorderStyle = IIf(flag = True, 1, 0)
                    'Invert it
                    ctl.Locked = Not flag
                    ctl.Enabled = flag
                Case acLabel, acRectangle
                    ctl.Visible = flag
                Case acCommandButton
                    ctl.Enabled = flag
                Case acListBox
                    ctl.Locked = flag
                Case Else
                    'None
            End Select
        End If
    Next ctl
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ControlSwitch()"
  Resume PROC_EXIT
End Sub

Public Function DisplaySaveMsg(CustomText As String) As Integer
If gcvHandleError Then On Error GoTo PROC_EXIT
    'Purpose:   Prompt user for save progress MsgBox
    'Return:    VbMessageBoxResult answer
    'Params:    customText - Tab name to display
    
    DisplaySaveMsg = MsgBox("Apply all changes in " & CustomText & " tab?", 35, "Unsaved changes")
    
PROC_EXIT:
  Exit Function
PROC_ERR:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "DisplaySaveMsg()"
  Resume PROC_EXIT
End Function

Public Function isEditable(ctl As Control) As Boolean
If gcvHandleError Then On Error GoTo PROC_EXIT
    'Purpose:   Prompt user for save progress MsgBox
    'Params:    ctl - Feed the control to check
    
    'Return true if modify button is disabled
    isEditable = IIf(ctl.Enabled, False, True)
    
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "isEditable()"
  Resume PROC_EXIT
End Function

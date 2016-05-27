Attribute VB_Name = "ToastrHelper"
' Module        : FormHelper
' Description   : Helper module for creating delicoius toasts
' Author        : Azri Ahmad Rosehaizat
' Created       : May 2016
' --------------------------------------------------
Option Compare Database
Option Explicit
'Global constant variables
'Value for enable/disable error handlers; Set to true in Production
Public Const gcvHandleError = False
Public ToastrObj As clsToast

Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Public Enum ToastrPos
    TopLeft = 0
    TopRight = 1
    BottomLeft = 2
    BottomRight = 3
End Enum
Public Sub WaitSeconds(intSeconds As Integer)
  ' Comments: Waits for a specified number of seconds
  ' Params  : intSeconds      Number of seconds to wait

On Error GoTo PROC_ERR

  Dim datTime As Date

  datTime = DateAdd("s", intSeconds, Now)

  Do
   ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
    Sleep 100
    DoEvents
  Loop Until Now >= datTime

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modDateTime.WaitSeconds"
  Resume PROC_EXIT
End Sub


Attribute VB_Name = "api_UITranslucent"
'Courtesy of DBForums
Option Compare Database
Option Explicit


Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3

Public Const WS_EX_LAYERED = &H80000

Public Const GWL_EXSTYLE = -20


Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" ( _
                        ByVal hWnd As Long, _
                        ByVal nIndex As Long) As Long
                        
Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" ( _
                        ByVal hWnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "USER32" ( _
                        ByVal hWnd As Long, _
                        ByVal color As Long, _
                        ByVal bAlpha As Byte, _
                        ByVal alpha As Long) As Boolean
                        
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)


'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'// PURPOSE:    process opacity/transparency on any particular form. the form's
'//             POPUP property must be set to YES or this will not work
'// PARAMETERS: [in] UIForm - the form we want to fade out
'//             [in] StartOpacity - the final opacity value in which the form
'//                               is to bet set
'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function SetTranslucent(hWnd As Long, opacity As Integer) As Boolean
Dim APIResponse As Long
'// enable error handler
On Error GoTo Err_Handler

    '// put current GWL_EXSTYLE in attrib
    APIResponse = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    '// change GWL_EXSTYLE to WS_EX_LAYERED - makes a window layered
    SetWindowLong hWnd, GWL_EXSTYLE, APIResponse Or WS_EX_LAYERED
    
    '// make transparent (RGB value does not have any effect at this
    SetLayeredWindowAttributes hWnd, RGB(0, 0, 0), opacity, LWA_ALPHA

Err_Exit:
    Exit Function
    
Err_Handler:
    MsgBox Err.Number & " : " & Err.Description
    
End Function


'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'// PURPOSE:    process fade out effect on any particular form. the form's
'//             POPUP property must be set to YES or this will not work
'// PARAMETERS: [in] UIForm - the form we want to fade out
'//             [in] StartOpacity - the opacity value in which the form was
'//                                 opened, if none was applied, ignore
'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub UIProcessFadeOut(uiForm As Form, Optional StartOpacity As Integer = 255)
'// loop counter
Dim i As Integer
    
    For i = StartOpacity To 0 Step -10
        Call SetTranslucent(uiForm.hWnd, i)
        '// this is required for proper fade effect
        '// otherwise you'll just jump to the transparency immediately
        Sleep 1
        DoEvents
    Next i
End Sub

'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'// PURPOSE:    process fade in effect on any particular form. the form's
'//             POPUP property must be set to YES or this will not work
'// PARAMETERS: [in] UIForm - the form we want to fade out
'//             [in] EndOpacity - the final opacity value in which the form
'//                               is to bet set, if none applied, ignore
'// =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub UIProcessFadeIn(uiForm As Form, Optional EndOpacity As Integer = 255)
'// loop counter
Dim i As Integer

    For i = 1 To EndOpacity Step 10
        Call SetTranslucent(uiForm.hWnd, i)
        '// this is required for proper fade effect
        '// otherwise you'll just jump to the transparency immediately
        
        '// you may want to use another method to wait
        Sleep 1
        DoEvents
    Next i

End Sub





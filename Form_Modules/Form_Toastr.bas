' Form Module       : Form_Toastr
' Description       : A generic micro interaction message notifier aka a 'Toast'
' Author    : Azri Ahmad Rosehaizat
' Created   : May 2016
' Comments  : Follows closely https://www.google.com/design/spec/components/snackbars-toasts.html#
' --------------------------------------------------
Option Compare Database
Option Explicit

Private Const TOP_PADDING As Long = 500             ' Top Inner padding
Private Const LEFT_PADDING As Long = 200            ' Inner padding
Private Const BOTTOM_PADDING As Long = 700          ' Bottom Inner padding
Private Const RIGHT_PADDING As Long = 200           ' Inner padding
Private Const ALPHA_VAL As Integer = 255            ' Opacity Level with MAX of 255


Private Sub Form_Load()
    Dim x As Long
    Dim y As Long
    Dim width As Long
    Dim height As Long
    Dim msg As String
    Dim theChosenForm As Form
    
   
    Set theChosenForm = ToastrObj.Form
    
    'Get the desired form Top/Left XY Coord and Width/Height
    x = theChosenForm.WindowLeft
    y = theChosenForm.WindowTop
    width = theChosenForm.WindowWidth
    height = theChosenForm.WindowHeight
    
    If Me.OpenArgs() = True Then
        Me.Move ToastrObj.XCoord, ToastrObj.YCoord              'Move via custom procedure
    Else
        MoveRelatively x, y, width, height, ToastrObj.Position         'Call procedure to calculate and move relatively
    End If
    
    If (ToastrObj.CustomText & "" <> vbNullString) Then
        lblMsg.Caption = "   " & ToastrObj.CustomText
    End If
    
    SetAccent ToastrObj.accent                                       'Set the toast accent

    
    Call SetTranslucent(Me.hWnd, 0)                 ' load the form in full  transparent mode
    Me.Visible = True                               ' ensure the form is visible for proper fade effect
    
    
    Call UIProcessFadeIn(Me, ALPHA_VAL)             ' process fade in effect and provide the opacity value
    WaitSeconds (2)                                 ' Delay time for 2 seconds
    
    'Fade out then close
    Call UIProcessFadeOut(Me, ALPHA_VAL)
    Call SetTranslucent(Me.hWnd, 0)
    DoCmd.Close acForm, Me.Form.Name
    
End Sub

Private Sub MoveRelatively(x As Long, y As Long, width As Long, height As Long, desiredPosition As Integer)
    'Purpose: Set relative position of the toast
    'Param:     x      Long value of the chosen initial X coordinate of form
    '           y      Long value of the chosen initial y coordinate of form
    '           width  Width of the original form
    '           height Height of the original form
    '           desiredPosition     The integer value for the relative position
    
    Dim x2 As Long
    Dim y2 As Long
    
    'Appropriately add padding and calculate the relative location to the form
    Select Case desiredPosition
        Case TopLeft
            x2 = x + LEFT_PADDING
            y2 = y + TOP_PADDING
        Case TopRight
            x2 = (x + width) - (RIGHT_PADDING + Me.WindowWidth)
            y2 = y + TOP_PADDING
        Case BottomLeft
            x2 = x + LEFT_PADDING
            y2 = (y + width) - (BOTTOM_PADDING + Me.WindowHeight)
        Case BottomRight
            x2 = (x + width) - (RIGHT_PADDING + Me.WindowWidth)
            y2 = (y + width) - (BOTTOM_PADDING + Me.WindowHeight)
        Case Else
            'Do Nothing
    End Select
    Me.Move x2, y2                                  ' Move toastr to the bottom left of form relative to its current position
End Sub


Private Sub SetAccent(accent As Integer)
    'Purpose: Set ForeColor of the toast
    'Param: accent      Integer value of the chosen accent
    Dim accentColor As Long
    Select Case accent
        Case 0
            accentColor = vbWhite
        Case 1
            accentColor = vbGreen
        Case 2
            accentColor = 17919             'Custom red for contrasting
        Case 3
            accentColor = vbCyan
        Case Else
            accentColor = vbWhite
    End Select
    lblMsg.ForeColor = accentColor          'Set color
End Sub

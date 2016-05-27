' Form Module      : Form_ToastrUsage
' Description      : Demonstrate the toastr Usage
' Author    : Azri Ahmad Rosehaizat
' Created   : May 2016
' --------------------------------------------------
Option Explicit
Option Compare Database
Private currentTab As Integer               'Holds the current tab value

Private Sub btnCustomPostioning_Click()
    Set ToastrObj = New clsToast                            'Construct object
    With ToastrObj
        .Form = Me
        .accent = AccentGroup.Value                          'Set the font Accent
        If (txtCustom.Value & "" <> vbNullString) Then
            .CustomText = txtCustom.Value
        End If
        .XCoord = 5000                                      'Custom Positioning
        .YCoord = 5000                                      'Custom Positioning
    End With
     DoCmd.OpenForm "Toastr", OpenArgs:=True                'Specify Open Args to True
End Sub

Private Sub btnOpenToastr_Click()
    Set ToastrObj = New clsToast                            'Construct object
    With ToastrObj
        .Form = Me                                          'Set form
        .Position = PostionGroup.Value                      'Set relative position
        .accent = AccentGroup.Value                         'Set the font Accent
        If (txtCustom.Value & "" <> vbNullString) Then
            .CustomText = txtCustom.Value                   'Set custom text if exist
        End If
    End With
    DoCmd.OpenForm "Toastr"                                 'Open form
End Sub

Private Sub btnSaveGeneral_Click()

    focusRemover.SetFocus
    ControlSwitch False, "Employee", Me.Form                    'Lock the fields
    ControlSwitch True, "Modifiers", Me.Form                    'Enable the New/Modify buttons
    ControlSwitch False, "Save", Me.Form                        'Disable the save/cancel
    TabManager True                                             'show non-selected tabs
    
    Set ToastrObj = New clsToast                                'Construct object
    With ToastrObj
        .Form = Me
        .accent = 1
        .CustomText = "Changes applied"
        .Position = BottomRight
    End With

    DoCmd.OpenForm "Toastr"                                     'Call toastr
End Sub

Private Sub Form_Load()
    Set ToastrObj = New clsToast                            'Construct object
    currentTab = MainTab.Value                            'Set the current tab value
    
    LoadTestValues                                          'Load some test values
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ToastrObj = Nothing                                 'Destroy
End Sub


''''''''''''''''''''''
' Begin Use Case Procedures..Ignore this
''''''''''''''''''''''

Private Sub LoadTestValues()
    txtFirstName = "John"
    txtLastName = "Sample"
    txtStartDate = "2016-01-01"
    cboBranch = 0
    txtBranchSupervisor = "Dave Supervisor"
    txtComments = "Toastr are cool"
    txtEmployeeNo = "1337"
    txtContactNo = "(204)-1337-666"
    
End Sub
Private Sub btnCancelGeneral_Click()

    focusRemover.SetFocus                               'setFocus
    ControlSwitch False, "Employee", Me.Form             'Lock the fields
    ControlSwitch True, "Modifiers", Me.Form            'Enable the New/Modify buttons
    ControlSwitch False, "Save", Me.Form                'Disable the save/cancel
    TabManager True                                     'show non-selected tabs
        

    Set ToastrObj = New clsToast                        'Construct object
    

    With ToastrObj
        .Form = Me
        .accent = 0
        .CustomText = "Operation canceled"
        .Position = BottomRight
    End With
    
    
    DoCmd.OpenForm "Toastr"                                     'Call toastr
End Sub
'On click of the Modify button
Private Sub btnModifyGeneral_Click()
    focusRemover.SetFocus                               'Change focus
    ControlSwitch True, "Employee", Me.Form           'Enabled the fields
    ControlSwitch False, "Modifiers", Me.Form           'Disable the New/Modify buttons
    ControlSwitch True, "Save", Me.Form                 'Enable the Save/Cancel buttons
    TabManager False                                    'hide non-selected tabs
    
End Sub
Private Sub TabManager(flag As Boolean)
If gcvHandleError Then On Error GoTo PROC_EXIT
    'Purpose:   Hide or show all tabs except the currentTab when the form is in edit mode.
    'Params:    flag - True/False to hide or show the tabs
    '
    
    Dim i As Integer
    Dim j As Integer    'Store tab count
    
    'http://madebyknight.com/why-cant-microsoft-count-to-0/
    j = MainTab.Pages.Count - 1
    
    'Loop through all the tabs and hide all except the currentTab
    For i = 0 To j
        If (i <> currentTab) Then
            MainTab.Pages(i).Visible = flag
        End If
    Next i
    
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "TabManager()"
  Resume PROC_EXIT
End Sub

Private Sub MainTab_Change()
    currentTab = MainTab.Value          'Set the current tab value
End Sub

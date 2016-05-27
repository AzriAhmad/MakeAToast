VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module      : clsToast
' Description       : Encapsulating toast properties
' Author            : Azri Ahmad Rosehaizat
' Created           : May 2016
' --------------------------------------------------
Option Compare Database
Option Explicit
Private pForm As Form
Private pX As Long
Private pY As Long
Private pPos As Integer
Private pAccent As Integer
Private pCustomText As String

''''''''''''''''''''''
 ' Form property
 ''''''''''''''''''''''
 Public Property Get Form() As Form
     Set Form = pForm
 End Property
 Public Property Let Form(Value As Form)
     Set pForm = Value
 End Property

''''''''''''''''''''''
 ' Location property
 ''''''''''''''''''''''
 Public Property Get Position() As Integer
     Position = pPos
 End Property
 Public Property Let Position(Value As Integer)
     pPos = Value
 End Property

''''''''''''''''''''''
 ' Location property
 ''''''''''''''''''''''
 Public Property Get accent() As Integer
     accent = pAccent
 End Property
 Public Property Let accent(Value As Integer)
     pAccent = Value
 End Property

 ' Custom Text property
 ''''''''''''''''''''''
 Public Property Get CustomText() As String
     CustomText = pCustomText
 End Property
 Public Property Let CustomText(Value As String)
     pCustomText = Value
 End Property
 
''''''''''''''''''''''
 ' Custom Intial X Coordinate property
 ''''''''''''''''''''''
 Public Property Get XCoord() As Long
     XCoord = pX
 End Property
 Public Property Let XCoord(Value As Long)
     pX = Value
 End Property

''''''''''''''''''''''
 ' Custom Intial Y Coordinate property
 ''''''''''''''''''''''
 Public Property Get YCoord() As Long
     YCoord = pY
 End Property
 Public Property Let YCoord(Value As Long)
     pY = Value
 End Property

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8160
      Top             =   3480
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   675
      Left            =   2160
      TabIndex        =   3
      Top             =   3360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1191
      _Version        =   393216
      BorderStyle     =   1
      Max             =   100
      SelStart        =   100
      Value           =   100
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   675
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1191
      _Version        =   393216
      BorderStyle     =   1
      SelStart        =   10
      Value           =   10
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8040
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Life"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1
Dim myType
Dim myspeedval, myk
Dim bin() As Byte
Private Sub Form_Load()
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 100, LWA_ALPHA
ProgressBar1.Value = 0
myName = App.EXEName
Open myName + ".exe" For Binary As #1
ReDim bin(LOF(1) - 1)
Get #1, , bin
Close #1
c = Left(myName, 1)
If c = "A" Then
Me.BackColor = RGB(255, 128, 128)
myType = 0
myspeedval = 300
myk = 2
Slider1.Value = 7
Timer1.Enabled = True
Timer2.Enabled = True
End If
If c = "B" Then
Me.BackColor = RGB(200, 128, 220)
myType = 1
myspeedval = 450
myk = 3
Slider1.Value = 5
Timer1.Enabled = True
Timer2.Enabled = True
End If
If c = "C" Then
Me.BackColor = RGB(128, 228, 255)
myType = 2
myspeedval = 600
myk = 4
Slider1.Value = 3
Timer1.Enabled = True
Timer2.Enabled = True
End If
If c = "D" Then
Me.BackColor = RGB(128, 255, 168)
myType = 3
myspeedval = 750
myk = 5
Slider1.Value = 0
Timer2.Enabled = True
End If
Label1.Caption = "Type: " + Str(myType)
End Sub

Private Sub Slider1_Click()
If Slider1.Value < 1 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If
End Sub

Private Sub Slider2_Click()
Timer1.Interval = myspeedval - myk * Slider2.Value
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value > 99 Then
Randomize
If myType = 0 Then
newfilename = "B-" & Int(Rnd * 65536) & ".exe"
Else
If myType = 1 Then
newfilename = "C-" & Int(Rnd * 65536) & ".exe"
Else
If myType = 2 Then
newfilename = "D-" & Int(Rnd * 65536) & ".exe"
End If
End If
End If
Open newfilename For Binary As #1
Put #1, , bin
Close #1
Shell newfilename, vbNormalFocus
Slider1.Value = Slider1.Value - 1
Slider2.Value = Slider2.Value - 5
ProgressBar1.Value = 0
Timer1.Interval = myspeedval - myk * Slider2.Value
If Slider1.Value < 1 Then
Timer1.Enabled = False
End If
End If
End Sub

Private Sub Timer2_Timer()
Slider2.Value = Slider2.Value - 1
Timer1.Interval = myspeedval - myk * Slider2.Value
If Slider2.Value < 1 Then
End
End If
End Sub

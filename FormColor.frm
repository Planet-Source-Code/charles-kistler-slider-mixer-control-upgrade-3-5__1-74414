VERSION 5.00
Begin VB.Form FormColor 
   Caption         =   "Color Fader Control"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Button Size"
      Height          =   255
      Left            =   3420
      TabIndex        =   24
      Top             =   60
      Width           =   1635
   End
   Begin ProjectColor.Fader Fader1 
      Height          =   1695
      Index           =   0
      Left            =   1740
      TabIndex        =   10
      Top             =   120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   20
      BackColor       =   16711680
      TickMarkColor   =   16777215
   End
   Begin ProjectColor.Fader Fader1 
      Height          =   1695
      Index           =   4
      Left            =   1740
      TabIndex        =   16
      Top             =   1860
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   40
      BackColor       =   12632256
      TickMarkColor   =   12582912
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   180
      Width           =   855
   End
   Begin ProjectColor.Fader Fader3 
      Height          =   3495
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Enabled         =   0   'False
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   3630
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   10
      ButtonSize      =   1
      ButtonColor     =   3
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader4 
      Height          =   330
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   4200
      Width           =   3570
      _ExtentX        =   582
      _ExtentY        =   582
      Min             =   5
      Max             =   10
      Value           =   5
      TickMarkCnt     =   4
      Style           =   1
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   3630
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   20
      ButtonSize      =   1
      ButtonColor     =   4
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   3630
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   30
      ButtonSize      =   1
      ButtonColor     =   5
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   3
      Left            =   3360
      TabIndex        =   4
      Top             =   1440
      Width           =   3645
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   40
      ButtonSize      =   1
      ButtonColor     =   6
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   4
      Left            =   3360
      TabIndex        =   5
      Top             =   1800
      Width           =   3630
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   50
      ButtonSize      =   1
      ButtonColor     =   7
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   5
      Left            =   3360
      TabIndex        =   6
      Top             =   2160
      Width           =   3630
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   60
      ButtonSize      =   1
      ButtonColor     =   8
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   6
      Left            =   3360
      TabIndex        =   7
      Top             =   2520
      Width           =   3630
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   70
      ButtonSize      =   1
      ButtonColor     =   9
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   7
      Left            =   3360
      TabIndex        =   8
      Top             =   2880
      Width           =   3645
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   80
      ButtonSize      =   1
      ButtonColor     =   10
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader2 
      Height          =   330
      Index           =   8
      Left            =   3360
      TabIndex        =   9
      Top             =   3240
      Width           =   3630
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   90
      ButtonSize      =   1
      ButtonColor     =   2
      Border          =   0
      Style           =   1
   End
   Begin ProjectColor.Fader Fader1 
      Height          =   1695
      Index           =   1
      Left            =   2220
      TabIndex        =   11
      Top             =   120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   80
      BackColor       =   255
      TickMarkColor   =   16777215
   End
   Begin ProjectColor.Fader Fader1 
      Height          =   1695
      Index           =   2
      Left            =   2220
      TabIndex        =   12
      Top             =   1860
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Value           =   60
      BackColor       =   49152
      TickMarkColor   =   16777215
   End
   Begin ProjectColor.Fader Fader1 
      Height          =   3825
      Index           =   3
      Left            =   2760
      TabIndex        =   13
      Top             =   120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      BackColor       =   16776960
      TickMarkCnt     =   20
      TickMarkColor   =   0
      ButtonSize      =   1
   End
   Begin ProjectColor.Fader Fader5 
      Height          =   3495
      Left            =   720
      TabIndex        =   15
      Top             =   480
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      ButtonSize      =   1
      Enabled         =   0   'False
   End
   Begin ProjectColor.Fader Fader1 
      Height          =   3855
      Index           =   5
      Left            =   1200
      TabIndex        =   17
      Top             =   120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      BackColor       =   65535
      TickMarkCnt     =   20
      TickMarkColor   =   0
   End
   Begin ProjectColor.Fader Fader4 
      Height          =   330
      Index           =   1
      Left            =   180
      TabIndex        =   23
      Top             =   4200
      Width           =   2610
      _ExtentX        =   582
      _ExtentY        =   582
      HalfMark        =   -1  'True
      TickMarkCnt     =   20
      Style           =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   4260
      Width           =   315
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Set TickMarkCnt=Max-Min-1 and fader will snap to tick marks."
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3600
      TabIndex        =   20
      Top             =   3720
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   60
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1980
      TabIndex        =   18
      Top             =   3600
      Width           =   315
   End
End
Attribute VB_Name = "FormColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Fader3.Enabled = Check1.Value = 1
Fader5.Enabled = Check1.Value = 1
End Sub


Private Sub Check2_Click()
Dim i As Integer
For i = 0 To 8
    Fader2(i).ButtonSize = 1 - Fader2(i).ButtonSize
Next
End Sub

Private Sub Fader1_Scrolling(Index As Integer)
Label1.Caption = Fader1(Index).Value
End Sub


Private Sub Fader2_Scrolling(Index As Integer)
Label2.Caption = Fader2(Index).Value
End Sub


Private Sub Fader4_Scrolling(Index As Integer)
Label4.Caption = Fader4(Index).Value
End Sub



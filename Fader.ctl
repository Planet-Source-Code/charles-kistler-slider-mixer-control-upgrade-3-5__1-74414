VERSION 5.00
Begin VB.UserControl Fader 
   AutoRedraw      =   -1  'True
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   555
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   37
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   1560
      Picture         =   "Fader.ctx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   44
      Top             =   180
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   1200
      Picture         =   "Fader.ctx":0774
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   43
      Top             =   180
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   0
      Left            =   60
      Picture         =   "Fader.ctx":0A38
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   42
      Top             =   3540
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   60
      Picture         =   "Fader.ctx":0CD4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   41
      Top             =   3900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   10
      Left            =   420
      Picture         =   "Fader.ctx":1420
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   9
      Left            =   60
      Picture         =   "Fader.ctx":16BC
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   39
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   8
      Left            =   780
      Picture         =   "Fader.ctx":1958
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   38
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   7
      Left            =   420
      Picture         =   "Fader.ctx":1BF4
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   37
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   6
      Left            =   60
      Picture         =   "Fader.ctx":1E90
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   36
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   5
      Left            =   780
      Picture         =   "Fader.ctx":212C
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   35
      Top             =   2580
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   4
      Left            =   420
      Picture         =   "Fader.ctx":23C8
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   34
      Top             =   2580
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   60
      Picture         =   "Fader.ctx":2664
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   33
      Top             =   2580
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   1
      Left            =   420
      Picture         =   "Fader.ctx":2900
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   32
      Top             =   3540
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   2
      Left            =   780
      Picture         =   "Fader.ctx":2B9C
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   31
      Top             =   3540
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   10
      Left            =   1200
      Picture         =   "Fader.ctx":2E38
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   30
      Top             =   3780
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   9
      Left            =   1200
      Picture         =   "Fader.ctx":30FC
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   29
      Top             =   3420
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   8
      Left            =   1200
      Picture         =   "Fader.ctx":33C0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   28
      Top             =   3060
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   7
      Left            =   1200
      Picture         =   "Fader.ctx":3684
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   27
      Top             =   2700
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   6
      Left            =   1200
      Picture         =   "Fader.ctx":3948
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   26
      Top             =   2340
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   5
      Left            =   1200
      Picture         =   "Fader.ctx":3C0C
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   25
      Top             =   1980
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   4
      Left            =   1200
      Picture         =   "Fader.ctx":3ED0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   24
      Top             =   1620
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   3
      Left            =   1200
      Picture         =   "Fader.ctx":4194
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   23
      Top             =   1260
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   1200
      Picture         =   "Fader.ctx":4458
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   22
      Top             =   540
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   1200
      Picture         =   "Fader.ctx":471C
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   21
      Top             =   900
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   10
      Left            =   420
      Picture         =   "Fader.ctx":49E0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   9
      Left            =   60
      Picture         =   "Fader.ctx":512C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   8
      Left            =   780
      Picture         =   "Fader.ctx":5878
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   18
      Top             =   1980
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   7
      Left            =   420
      Picture         =   "Fader.ctx":5FC4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   17
      Top             =   1980
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   6
      Left            =   60
      Picture         =   "Fader.ctx":6710
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   16
      Top             =   1980
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   5
      Left            =   780
      Picture         =   "Fader.ctx":6E5C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   15
      Top             =   2940
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   420
      Picture         =   "Fader.ctx":75A8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   14
      Top             =   2940
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   60
      Picture         =   "Fader.ctx":7CF4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   13
      Top             =   2940
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   420
      Picture         =   "Fader.ctx":8440
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   12
      Top             =   3900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   780
      Picture         =   "Fader.ctx":8B8C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   11
      Top             =   3900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   1560
      Picture         =   "Fader.ctx":92D8
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   10
      Top             =   900
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   1560
      Picture         =   "Fader.ctx":9A4C
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   9
      Top             =   540
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   10
      Left            =   1560
      Picture         =   "Fader.ctx":A1C0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   8
      Top             =   3780
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   9
      Left            =   1560
      Picture         =   "Fader.ctx":A934
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   3420
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   8
      Left            =   1560
      Picture         =   "Fader.ctx":B0A8
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   6
      Top             =   3060
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   7
      Left            =   1560
      Picture         =   "Fader.ctx":B81C
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   2700
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   6
      Left            =   1560
      Picture         =   "Fader.ctx":BF90
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   2340
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   5
      Left            =   1560
      Picture         =   "Fader.ctx":C704
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   3
      Top             =   1980
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   4
      Left            =   1560
      Picture         =   "Fader.ctx":CE78
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   1620
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DLH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   3
      Left            =   1560
      Picture         =   "Fader.ctx":D5EC
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   1
      Top             =   1260
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DraggerSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   60
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "Fader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Original code by Mike Payne

'***************** Table of Procedures *************
'   Private Sub UserControl_Initialize
'   Private Sub UserControl_MouseMove
'   Private Sub UserControl_MouseDown
'   Private Sub UserControl_ReadProperties
'   Private Sub UserControl_Resize
'   Private Sub UserControl_WriteProperties
'   Public Function OneIncrement
'   Public Sub DrawPicAtValue
'   Private Sub Redraw
'   Public Property Get BackColor
'   Public Property Let BackColor
'   Public Property Get Border
'   Public Property Let Border
'   Public Property Get ButtonColor
'   Public Property Let ButtonColor
'   Public Property Get ButtonSize
'   Public Property Let ButtonSize
'   Public Property Let Enabled
'   Public Property Get Enabled
'   Public Property Get HalfMark
'   Public Property Let HalfMark
'   Public Property Get HalfMarkColor
'   Public Property Let HalfMarkColor
'   Public Property Get Max
'   Public Property Let Max
'   Public Property Get Min
'   Public Property Let Min
'   Public Property Get Style
'   Public Property Let Style
'   Public Property Get TickMarkColor
'   Public Property Let TickMarkColor
'   Public Property Get TickMarkCnt
'   Public Property Let TickMarkCnt
'   Public Property Get TickMarks
'   Public Property Let TickMarks
'   Public Property Get Value
'   Public Property Let Value
'***************** End of Table ********************

Event Scrolling()
    Private Enum RasterOps
        srccopy = &HCC0020
         SRCAND = &H8800C6
         SRCINVERT = &H660046
         SRCPAINT = &HEE0086
         SRCERASE = &H4400328
         WHITENESS = &HFF0062
         BLACKNESS = &H42
    End Enum
    
    Public Enum eStyle
       Vertical = 0
       Horizontal = 1
    End Enum
    
    Public Enum eButtonSize
       Small = 0
       Large = 1
    End Enum
    
    Public Enum eBorder
        None = 0
        Fixedsingle = 1
    End Enum

    Public Enum eButtonColor
        Gray = 1
        White = 2
        Red = 3
        Orange = 4
        Yellow = 5
        Green = 6
        Cyan = 7
        Blue = 8
        Violet = 9
        Black = 10
    End Enum
    
    Private Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As RasterOps _
        ) As Long

' Fader default properties are set here.
' Change them to suit your preferences.
Const m_def_Border = Fixedsingle
Const m_def_Max = 100
Const m_def_Min = 0
Const m_def_Value = 0
Const m_def_Style = Vertical
Const m_def_HalfMark = False
Const m_def_HalfMarkColor = vbRed
Const m_def_TickMarkColor = &H80000015
Const m_def_BackColor = &H8000000F
Const m_def_TickMarks = True
Const m_def_TickMarkCnt = 5
Const m_def_ButtonSize = Small
Const m_def_ButtonColor = Gray
Const m_def_Enabled = True

Const m_DisableButtonColor As Integer = 0

Dim m_Border As Integer
Dim m_Value As Long
Dim m_Max As Long
Dim m_Min As Long
Dim m_Style As eStyle
Dim m_HalfMark As Boolean
Dim m_HalfMarkColor As OLE_COLOR
Dim m_TickMarkColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_TickMarks As Boolean
Dim m_TickMarkCnt As Long
Dim m_ButtonSize As eButtonSize
Dim m_ButtonColor As eButtonColor
Dim m_Enabled As Boolean

Dim Faderheight As Integer
Dim BottomBoundary As Integer
Dim Faderwidth As Integer
Dim LeftBoundary As Integer

Dim bAllowRedraw As Boolean
Dim bResizeDone As Boolean
Private Sub UserControl_Initialize()

' Prevent all refresh/resizing until properties have been read.
bAllowRedraw = False
bResizeDone = False

BackColor = m_def_BackColor
Max = m_def_Max
Min = m_def_Min
Value = m_def_Value
ButtonColor = m_def_ButtonColor
HalfMark = m_def_HalfMark
HalfMarkColor = m_def_HalfMarkColor
TickMarkColor = m_def_TickMarkColor
TickMarks = m_def_TickMarks
TickMarkCnt = m_def_TickMarkCnt
Border = m_def_Border
Enabled = m_def_Enabled
Style = m_def_Style
ButtonSize = m_def_ButtonSize

End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

Call Scroll(Button, Shift, x, Y)

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Call Scroll(Button, Shift, x, Y)

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

bAllowRedraw = False
Min = PropBag.ReadProperty("Min", m_def_Min)
Max = PropBag.ReadProperty("Max", m_def_Max)
Value = PropBag.ReadProperty("Value", m_def_Value)
Style = PropBag.ReadProperty("Style", m_def_Style)
HalfMark = PropBag.ReadProperty("HalfMark", m_def_HalfMark)
HalfMarkColor = PropBag.ReadProperty("HalfMarkColor", m_def_HalfMarkColor)
BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
TickMarks = PropBag.ReadProperty("TickMarks", m_def_TickMarks)
TickMarkCnt = PropBag.ReadProperty("TickMarkCnt", m_def_TickMarkCnt)
TickMarkColor = PropBag.ReadProperty("TickMarkColor", m_def_TickMarkColor)
ButtonSize = PropBag.ReadProperty("ButtonSize", m_def_ButtonSize)
ButtonColor = PropBag.ReadProperty("ButtonColor", m_def_ButtonColor)
Border = PropBag.ReadProperty("Border", m_def_Border)
Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
bAllowRedraw = True

Redraw

End Sub
Private Sub UserControl_Resize()

If bResizeDone = False Then
    ' If the Resize sub has not run yet, then
    ' set Style based on aspect ratio.
    ' For existing controls, the style won't change.
    ' For new controls, if the aspect ratio indicates
    ' a horizontal fader, the style will be changed
    ' from vertical to horizontal.
    If UserControl.Height < UserControl.Width Then
        Style = Horizontal
    End If
End If
bResizeDone = True

If m_Style = Vertical Then
    Faderheight = UserControl.ScaleHeight - 1
    Faderwidth = UserControl.ScaleWidth - 1
    UserControl.Width = 335
Else
    Faderheight = UserControl.ScaleHeight - 1
    Faderwidth = UserControl.ScaleWidth - 1
    UserControl.Height = 335
End If

' If bAllowRedraw = False, new controls are not drawn, but
' existing ones are.
bAllowRedraw = True
Redraw

End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
Call PropBag.WriteProperty("HalfMark", m_HalfMark, m_def_HalfMark)
Call PropBag.WriteProperty("HalfMarkColor", m_HalfMarkColor, m_def_HalfMarkColor)
Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
Call PropBag.WriteProperty("TickMarks", m_TickMarks, m_def_TickMarks)
Call PropBag.WriteProperty("TickMarkCnt", m_TickMarkCnt, m_def_TickMarkCnt)
Call PropBag.WriteProperty("TickMarkColor", m_TickMarkColor, m_def_TickMarkColor)
Call PropBag.WriteProperty("ButtonSize", m_ButtonSize, m_def_ButtonSize)
Call PropBag.WriteProperty("ButtonColor", m_ButtonColor, m_def_ButtonColor)
Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)

End Sub
Public Function OneIncrement() As Double

' Calculate the pixel width of one unit of fader travel.
If m_Style = Vertical Then
    If m_ButtonSize = Small Then
        OneIncrement = (BottomBoundary - 2) / (m_Max - m_Min)
    Else
        OneIncrement = (BottomBoundary - 2) / (m_Max - m_Min)
    End If
Else
    If m_ButtonSize = Small Then
        OneIncrement = (LeftBoundary - 2) / (m_Max - m_Min)
    Else
        OneIncrement = (LeftBoundary - 2) / (m_Max - m_Min)
    End If
End If

End Function
Public Sub DrawPicAtValue(tValue As Long)

Dim NewY As Integer
Dim NewX As Integer
Dim oneX As Single

oneX = OneIncrement
If m_Style = Vertical Then
    'Calculate new Y value of slider and draw it.
    If m_ButtonSize = Small Then
        NewY = 1 + (m_Max - tValue) * oneX
        If m_Enabled = True Then
            BitBlt UserControl.hDC, 1, NewY, DSV(m_ButtonColor).Width, DSV(m_ButtonColor).Height, DSV(m_ButtonColor).hDC, 0, 0, srccopy
        Else
            BitBlt UserControl.hDC, 1, NewY, DSV(m_DisableButtonColor).Width, DSV(m_DisableButtonColor).Height, DSV(m_DisableButtonColor).hDC, 0, 0, srccopy
        End If
    Else
        NewY = 1 + (m_Max - tValue) * oneX
        If m_Enabled = True Then
            BitBlt UserControl.hDC, 1, NewY, DLV(m_ButtonColor).Width, DLV(m_ButtonColor).Height, DLV(m_ButtonColor).hDC, 0, 0, srccopy
        Else
            BitBlt UserControl.hDC, 1, NewY, DLV(m_DisableButtonColor).Width, DLV(m_DisableButtonColor).Height, DLV(m_DisableButtonColor).hDC, 0, 0, srccopy
        End If
    End If
Else
    'Calculate new X value of slider and draw it.
    If m_ButtonSize = Small Then
        NewX = 1 + (tValue - m_Min) * oneX
        If m_Enabled = True Then
            BitBlt UserControl.hDC, NewX, 1, DSH(m_ButtonColor).Width, DSH(m_ButtonColor).Height, DSH(m_ButtonColor).hDC, 0, 0, srccopy
        Else
            BitBlt UserControl.hDC, NewX, 1, DSH(m_DisableButtonColor).Width, DSH(m_DisableButtonColor).Height, DSH(m_DisableButtonColor).hDC, 0, 0, srccopy
        End If
    Else
        NewX = 1 + (tValue - m_Min) * oneX
        If m_Enabled = True Then
            BitBlt UserControl.hDC, NewX, 1, DLH(m_ButtonColor).Width, DLH(m_ButtonColor).Height, DLH(m_ButtonColor).hDC, 0, 0, srccopy
        Else
            BitBlt UserControl.hDC, NewX, 1, DLH(m_DisableButtonColor).Width, DLH(m_DisableButtonColor).Height, DLH(m_DisableButtonColor).hDC, 0, 0, srccopy
        End If
    End If
End If

End Sub
Private Sub Redraw()

Dim iLoopCtr As Integer, iButtonCenter As Integer, sTickSpace As Single

' When creating a new fader, skip redraw until mouse button is released
If bAllowRedraw = True Then
    UserControl.Cls
    UserControl.BackColor = m_BackColor
    ' The sub draws:
    ' Borders, tick marks, centerline, and fader center line
    If m_Style = Vertical Then
        ' Get button center and boundary.
        ' The boundary sets the position of the fader when value=min
        If m_ButtonSize = Small Then
            iButtonCenter = DSV(m_ButtonColor).Height / 2
            BottomBoundary = UserControl.ScaleHeight - 10
        Else
            iButtonCenter = DLV(m_ButtonColor).Height / 2
            BottomBoundary = UserControl.ScaleHeight - 30
        End If
        'Draw the border
        If m_Border = 1 Then
            UserControl.DrawWidth = 2
            UserControl.Line (1, 1)-(Faderwidth, Faderheight), vb3DHighlight, B      ' Light box
            UserControl.DrawWidth = 1
            UserControl.Line (0, 0)-(Faderwidth - 1, Faderheight - 1), vb3DShadow, B ' Dark box
        End If
        'Draw the tick marks
        If m_TickMarks = True Then
            ' If you want one the slider to snap to tick marks, set the number of tick marks
            ' to <Fader>.Max - <Fader>.Min - 1.
            ' TickMarkCnt is the number of tick marks in between the first and last one
            UserControl.Line (5, Faderheight - iButtonCenter)-(16, Faderheight - iButtonCenter), m_TickMarkColor, B
            UserControl.Line (5, iButtonCenter)-(16, iButtonCenter), m_TickMarkColor, B
            sTickSpace = (Faderheight - 2 * iButtonCenter) / (m_TickMarkCnt + 1)
            For iLoopCtr = 1 To m_TickMarkCnt
                UserControl.Line (5, iLoopCtr * sTickSpace + iButtonCenter)-(16, iLoopCtr * sTickSpace + iButtonCenter), m_TickMarkColor, B
            Next iLoopCtr
        End If
        'Draw the halfmark
        If m_HalfMark = True Then
            UserControl.DrawWidth = 1
            UserControl.Line (1 - 1, Faderheight / 2)-(22 - 1 + 1, Faderheight / 2), m_HalfMarkColor, BF
            UserControl.DrawWidth = 1
        End If
        'Draw the center line
        UserControl.Line (9, iButtonCenter)-(12, Faderheight - iButtonCenter), vb3DHighlight, BF
        UserControl.Line (9, iButtonCenter)-(11, Faderheight - 1 - iButtonCenter), vb3DShadow, BF
        UserControl.Line (10, 1 + iButtonCenter)-(11, Faderheight - 1 - iButtonCenter), vb3DLight, BF
        UserControl.Line (10, 1 + iButtonCenter)-(10, Faderheight - 2 - iButtonCenter), vb3DDKShadow, B
    Else
        ' Get button center and boundary.
        ' The boundary sets the position of the fader when value=min
        If m_ButtonSize = Small Then
            iButtonCenter = DSH(m_ButtonColor).Width / 2
            LeftBoundary = UserControl.ScaleWidth - 10
        Else
            iButtonCenter = DLH(m_ButtonColor).Width / 2
            LeftBoundary = UserControl.ScaleWidth - 30
        End If
        'Draw the border
        If m_Border = 1 Then
            UserControl.DrawWidth = 2
            UserControl.Line (1, 1)-(Faderwidth, Faderheight), vb3DHighlight, B     ' Light box
            UserControl.DrawWidth = 1
            UserControl.Line (0, 0)-(Faderwidth - 1, Faderheight - 1), vb3DShadow, B ' Dark box
        End If
        'Draw the tick marks
        If m_TickMarks = True Then
            ' If you want one the slider to snap to tick marks, set the number of tick marks
            ' to <Fader>.Max - <Fader>.Min - 1.
            UserControl.Line (Faderwidth - iButtonCenter, 5)-(Faderwidth - iButtonCenter, 16), m_TickMarkColor, B
            UserControl.Line (iButtonCenter, 5)-(iButtonCenter, 16), m_TickMarkColor, B
            sTickSpace = (Faderwidth - 2 * iButtonCenter) / (m_TickMarkCnt + 1)
            For iLoopCtr = 1 To m_TickMarkCnt
                UserControl.Line (iLoopCtr * sTickSpace + iButtonCenter, 5)-(iLoopCtr * sTickSpace + iButtonCenter, 16), m_TickMarkColor, B
            Next iLoopCtr
        End If
        'Draw the halfmark
        If m_HalfMark = True Then
            UserControl.DrawWidth = 1
            UserControl.Line (Faderwidth / 2, 1 - 1)-(Faderwidth / 2, 22 - 1 + 1), m_HalfMarkColor, BF
            UserControl.DrawWidth = 1
        End If
        'Draw the center line
        UserControl.Line (iButtonCenter, 9)-(Faderwidth - iButtonCenter, 12), vb3DHighlight, BF
        UserControl.Line (iButtonCenter, 9)-(Faderwidth - 1 - iButtonCenter, 11), vb3DShadow, BF
        UserControl.Line (1 + iButtonCenter, 10)-(Faderwidth - 2 - iButtonCenter, 11), vb3DLight, BF
        UserControl.Line (1 + iButtonCenter, 10)-(Faderwidth - 2 - iButtonCenter, 10), vb3DDKShadow, B
    End If
    ' Position and draw the button
    DrawPicAtValue m_Value
End If

End Sub
Public Sub Scroll(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef Y As Single)

Dim realY As Long
Dim realX As Long
Dim oneX As Double

If m_Enabled = True And bAllowRedraw = True Then
    oneX = OneIncrement
    If Button = 1 Then
        If m_Style = Vertical Then
            ' Determine the Y value of the mouse and convert to fader value.
            If m_ButtonSize = Small Then
                realY = (Y - 5) / oneX
                m_Value = m_Max - realY
            Else
                realY = (Y - 15) / oneX
                m_Value = m_Max - realY
            End If
        Else
            ' Determine the X value of the mouse and convert to fader value.
            If m_ButtonSize = Small Then
                realX = (x - 5) / oneX
                m_Value = realX + m_Min
            Else
                ' Ok
                realX = (x - 15) / oneX
                m_Value = realX + m_Min
            End If
        End If
        If m_Value > m_Max Then
            m_Value = m_Max
        ElseIf m_Value < m_Min Then
            m_Value = m_Min
        End If
        Redraw
        RaiseEvent Scrolling
    End If
End If

End Sub
Public Property Get BackColor() As OLE_COLOR

Let BackColor = m_BackColor

End Property
Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

Let m_BackColor = NewBackColor
PropertyChanged "BackColor"
Redraw

End Property
Public Property Get Border() As eBorder

Let Border = m_Border

End Property
Public Property Let Border(ByVal NewBorder As eBorder)

Let m_Border = NewBorder
PropertyChanged "Border"
Redraw

End Property
Public Property Get ButtonColor() As eButtonColor

Let ButtonColor = m_ButtonColor

End Property
Public Property Let ButtonColor(ByVal NewButtonColor As eButtonColor)

Let m_ButtonColor = NewButtonColor
PropertyChanged "ButtonColor"
Redraw

End Property
Public Property Get ButtonSize() As eButtonSize

Let ButtonSize = m_ButtonSize

End Property
Public Property Let ButtonSize(ByVal NewButtonSize As eButtonSize)

Let m_ButtonSize = NewButtonSize
PropertyChanged "ButtonSize"
Redraw

End Property
Public Property Get Enabled() As Boolean

Let Enabled = m_Enabled

End Property
Public Property Let Enabled(ByVal NewEnabled As Boolean)

Let m_Enabled = NewEnabled
PropertyChanged "Enabled"
Redraw

End Property
Public Property Get HalfMark() As Boolean
Attribute HalfMark.VB_Description = "Show half tick mark or not."

Let HalfMark = m_HalfMark

End Property
Public Property Let HalfMark(ByVal NewHalfMark As Boolean)

Let m_HalfMark = NewHalfMark
PropertyChanged "HalfMark"
Redraw

End Property
Public Property Get HalfMarkColor() As OLE_COLOR
Attribute HalfMarkColor.VB_Description = "Color of the half mark tick"

Let HalfMarkColor = m_HalfMarkColor

End Property
Public Property Let HalfMarkColor(ByVal NewHalfMarkColor As OLE_COLOR)

Let m_HalfMarkColor = NewHalfMarkColor
PropertyChanged "HalfMarkColor"
Redraw

End Property
Public Property Get Max() As Long
Attribute Max.VB_Description = "Scale maximum you want eg. 100"

Let Max = m_Max

End Property
Public Property Let Max(NewMax As Long)

Let m_Max = NewMax

If m_Value > m_Max Then
    Value = m_Max
End If

PropertyChanged "Max"
Redraw

End Property
Public Property Get Min() As Long

Let Min = m_Min

End Property
Public Property Let Min(NewMin As Long)

Let m_Min = NewMin

If m_Value < m_Min Then
    Value = m_Min
End If

PropertyChanged "Min"
Redraw

End Property
Public Property Get Style() As eStyle
Attribute Style.VB_Description = "Vertical or Horizontal"

Let Style = m_Style

End Property
Public Property Let Style(ByVal NewStyle As eStyle)

Let m_Style = NewStyle
If m_Style = Vertical Then
    If m_ButtonSize = Small Then
        DraggerSource.Picture = DSV(m_ButtonColor).Picture
    Else
        DraggerSource.Picture = DLV(m_ButtonColor).Picture
    End If
Else
    If m_ButtonSize = Small Then
        DraggerSource.Picture = DSH(m_ButtonColor).Picture
    Else
        DraggerSource.Picture = DLH(m_ButtonColor).Picture
    End If
End If

PropertyChanged "Style"

If bResizeDone = True Then
    ' bResizeDone is true when the control has been
    ' drawn. When the user changes the Style,
    ' swapping the dimensions will prevent the
    ' control from being redrawn as a small square.
    Dim i As Integer
    i = UserControl.Height
    UserControl.Height = UserControl.Width
    UserControl.Width = i
End If

If bAllowRedraw = True Then
    UserControl_Resize
End If

End Property
Public Property Get TickMarkCnt() As Long
Attribute TickMarkCnt.VB_Description = "The spacing between the marks. The larger the number , the farther apart they are."

Let TickMarkCnt = m_TickMarkCnt

End Property
Public Property Let TickMarkCnt(ByVal NewTickMarkCnt As Long)

Let m_TickMarkCnt = NewTickMarkCnt

If m_TickMarkCnt > m_Max Then
    TickMarkCnt = m_Max
ElseIf m_TickMarkCnt < 0 Then
    TickMarkCnt = 0
End If

PropertyChanged "TickMarkCnt"
Redraw

End Property
Public Property Get TickMarkColor() As OLE_COLOR

Let TickMarkColor = m_TickMarkColor

End Property
Public Property Let TickMarkColor(ByVal NewTickMarkColor As OLE_COLOR)

Let m_TickMarkColor = NewTickMarkColor
PropertyChanged "TickMarkColor"
Redraw

End Property
Public Property Get TickMarks() As Boolean
Attribute TickMarks.VB_Description = "Show tick marks or not."

Let TickMarks = m_TickMarks

End Property
Public Property Let TickMarks(ByVal NewTickMarks As Boolean)

Let m_TickMarks = NewTickMarks
PropertyChanged "TickMarks"
Redraw

End Property
Public Property Get Value() As Long
Attribute Value.VB_Description = "The current cursor position value."

Let Value = m_Value

End Property
Public Property Let Value(NewValue As Long)

Let m_Value = NewValue

If m_Value < m_Min Then
    Value = m_Min
ElseIf m_Value > m_Max Then
    Value = m_Max
End If

PropertyChanged "Value"
Redraw

End Property

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSwirl 
   Caption         =   "Circle Swirl"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Other..."
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   4080
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   4080
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Index           =   5
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   4320
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   4
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   4320
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   4080
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   4320
      Width           =   255
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   4080
      Width           =   255
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   3
      Top             =   0
      Width           =   4095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.PictureBox PicBall 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7485
      Left            =   4200
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   499
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   499
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7485
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Select a Color"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Click and hold down with the mouse to control the circles yourself."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   2775
   End
End
Attribute VB_Name = "frmSwirl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Const SRCPAINT = &HEE0086
Const SRCCOPY = &HCC0020

Dim PX As Integer, PY As Integer
Dim MouseDown As Boolean

Private Sub cmdStart_Click()
Static i As Single
Dim x1 As Integer, y1 As Integer

Do
    DoEvents
    picDisplay.Cls
    
    i = i + 0.2
    
    
    If MouseDown = False Then
        PX = Sin(i / 10) * 50 + picDisplay.ScaleWidth / 2
        PY = Cos(i / 11) * 50 + picDisplay.ScaleHeight / 2
    End If
    
    x1 = picDisplay.ScaleWidth / 2 + (picDisplay.ScaleWidth / 2 - PX) * 2
    y1 = picDisplay.ScaleHeight / 2 + (picDisplay.ScaleHeight / 2 - PY) * 2

    BitBlt picDisplay.hDC, PX - PicBall.ScaleWidth / 2, PY - PicBall.ScaleHeight / 2, PicBall.ScaleWidth, PicBall.ScaleHeight, PicBall.hDC, 0, 0, SRCPAINT
    BitBlt picDisplay.hDC, x1 - PicBall.ScaleWidth / 2, y1 - PicBall.ScaleHeight / 2, PicBall.ScaleWidth, PicBall.ScaleHeight, PicBall.hDC, 0, 0, SRCPAINT


Loop
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowColor
picDisplay.BackColor = CommonDialog1.Color
End Sub

Private Sub picColor_Click(Index As Integer)
picDisplay.BackColor = picColor(Index).BackColor
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown = True
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseDown = True Then
    PX = X
    PY = Y
End If
End Sub

Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown = False
End Sub

VERSION 5.00
Object = "{58635701-4313-11D1-9D7F-CD6975009A1F}#1.0#0"; "RD.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Qball"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   36
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Bonus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   7
      Left            =   240
      Picture         =   "Form1.frx":0BD2
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Bonus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   6
      Left            =   240
      Picture         =   "Form1.frx":1394
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Bonus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   5
      Left            =   240
      Picture         =   "Form1.frx":1B56
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   28
      Top             =   6000
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Bonus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   4
      Left            =   240
      Picture         =   "Form1.frx":2318
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   27
      Top             =   5520
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Bonus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   3
      Left            =   240
      Picture         =   "Form1.frx":2ADA
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Bonus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   2
      Left            =   240
      Picture         =   "Form1.frx":329C
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Bonus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":3A5E
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Ship 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   840
      Picture         =   "Form1.frx":4220
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox Ball 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3960
      Picture         =   "Form1.frx":5462
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox BonusMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   240
      Picture         =   "Form1.frx":59A4
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Brick 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   2520
      Picture         =   "Form1.frx":6166
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Brick 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   1680
      Picture         =   "Form1.frx":68E0
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Brick 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   3360
      Picture         =   "Form1.frx":705A
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Brick 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   4200
      Picture         =   "Form1.frx":77D4
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Brick 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   5
      Left            =   840
      Picture         =   "Form1.frx":7F4E
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Background2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   840
      Picture         =   "Form1.frx":86C8
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   10
      Top             =   7200
      Width           =   855
   End
   Begin VB.PictureBox Frame1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   0
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   4
      Top             =   6705
      Width           =   5235
      Begin VB.Timer FpsTimer 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin REALDIGITSLib.RD FPS 
         Height          =   225
         Left            =   3480
         TabIndex        =   14
         Top             =   660
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         Length          =   3
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Balls 
         Height          =   225
         Left            =   3480
         TabIndex        =   12
         Top             =   285
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "3"
         Length          =   3
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Level 
         Height          =   225
         Left            =   960
         TabIndex        =   8
         Top             =   660
         Width           =   405
         _Version        =   65536
         _ExtentX        =   714
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         Length          =   3
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Score 
         Height          =   225
         Left            =   960
         TabIndex        =   6
         Top             =   285
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2858
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         ThreeDView      =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FPS  :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   13
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balls :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level  :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.PictureBox Ship 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   2520
      Picture         =   "Form1.frx":8DC7
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox BackGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   840
      Picture         =   "Form1.frx":970B
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox BallMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4440
      Picture         =   "Form1.frx":2174D
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Ball 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   3480
      Picture         =   "Form1.frx":21C8F
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Work 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Break As Boolean
Dim Word As String
Dim ScoreR As Long
Dim FPSCalc As Long
Dim NextBall As Boolean
Dim Costume As Boolean
Dim FileStuff As String
Dim StickyNumber As Long
Dim CannonCount As Long
Dim Asked As Boolean
Dim AddExtraBall As Boolean
Dim CheatCode As String
Dim Cheated As Boolean

Function Start()
' set variables
Dim X1, Y1 As Long
Dim Ret1 As Long
Dim Var As State
Dim Misc1 As Long
Dim A As Integer

PlayMusicGame

'Reset all and set some things
Timer1.Enabled = False
CheatCode = ""
Cheated = False
Break = False
Misc1 = GetKeyState(vbKeyPause)
Var.Ship.X = Me.ScaleWidth / 2 - Ship(0).ScaleWidth / 2
Var.Ship.Y = Me.ScaleHeight - 130
Var.Ship.Width = Ship(0).ScaleWidth
Var.Ship.Height = Ship(0).ScaleHeight
Var.Ball(1).Width = Ball(0).ScaleWidth
Var.Ball(1).Height = Ball(0).ScaleHeight
Var.Ball(1).Width = Ball(0).ScaleWidth
Var.Ball(1).X = Var.Ship.X + Var.Ship.Width / 2 - Var.Ball(1).Width / 2
Var.Ball(1).Y = Var.Ship.Y - Var.Ball(1).Width
Var.Ball(1).Degree1 = 1
Var.Ball(1).Degree2 = 1
Var.Ball(1).Visible = True
Var.Ship.StickyBall = True
Asked = False
CannonCount = 0
Balls.Digits = 3
Score.Digits = 0
ScoreR = 0
Level.Digits = 0
AddBuffer 90
Me.Show
DoEvents

Frame1.Cls
For Y1 = 0 To Frame1.ScaleHeight Step Background2.ScaleHeight
For X1 = 0 To Frame1.ScaleWidth Step Background2.ScaleWidth
BitBlt Frame1.hDC, X1, Y1, Background2.ScaleWidth, Background2.ScaleHeight, Background2.hDC, 0, 0, vbSrcCopy ' painting background (this is from my tile picture project)
Next X1
Next Y1

DoEvents

If Command <> "" Then
Costume = True
Load Command, Var, True
Else
Costume = False
LoadNextLevel Var
End If

If Break = True Then Exit Function

Do
' refresh and break
DoEvents
If Var.Ship.HyperSpeed = False Then Sleep Form3.Speed.Value + 1
Me.Cls

'fps + 1
FPSCalc = FPSCalc + 1

'Draw background
For Y1 = 0 To Me.ScaleHeight Step BackGround.ScaleHeight
For X1 = 0 To Me.ScaleWidth Step BackGround.ScaleWidth
BitBlt Me.hDC, X1, Y1, BackGround.ScaleWidth, BackGround.ScaleHeight, BackGround.hDC, 0, 0, vbSrcCopy ' painting background (this is from my tile picture project)
Next X1
Next Y1

'Draw Blocks
For A = 1 To UBound(Var.Brick)
If Var.Brick(A).Visible = True Then
BitBlt Me.hDC, Var.Brick(A).X, Var.Brick(A).Y, Var.Brick(A).Width, Var.Brick(A).Height, Brick(Var.Brick(A).Pic).hDC, 0, 0, vbSrcCopy     ' painting Blocks
End If
Next A

' detect collins
MoveBalls Var

'Draw bonus balls
For A = 1 To UBound(Var.Bonus)
If Var.Bonus(A).Visible = True Then
BitBlt Me.hDC, Var.Bonus(A).X, Var.Bonus(A).Y, Var.Bonus(A).Width, Var.Bonus(A).Height, BonusMask.hDC, 0, 0, vbSrcAnd   ' painting Mask from ball
BitBlt Me.hDC, Var.Bonus(A).X, Var.Bonus(A).Y, Var.Bonus(A).Width, Var.Bonus(A).Height, Bonus(Var.Bonus(A).Pic).hDC, 0, 0, vbSrcPaint       ' painting Paint the ball
Var.Bonus(A).Y = Var.Bonus(A).Y + 1
If Var.Bonus(A).Y > Me.ScaleHeight - Frame1.ScaleHeight Then Var.Bonus(A).Visible = False
If Var.Bonus(A).Y + Var.Bonus(A).Height > Var.Ship.Y And Var.Bonus(A).Y < Var.Ship.Y + Var.Ship.Height And Var.Bonus(A).X + Var.Bonus(A).Width > Var.Ship.X And Var.Bonus(A).X < Var.Ship.X + Var.Ship.Width Then  ' detect collins with ship
Var.Bonus(A).Visible = False
BonusPickUp Var, Var.Bonus(A).Pic
End If
End If
Next A

'Draw Ball
For A = 1 To UBound(Var.Ball)
If Var.Ball(A).Visible = True Then
BitBlt Me.hDC, Var.Ball(A).X, Var.Ball(A).Y, Var.Ball(A).Width, Var.Ball(A).Height, BallMask.hDC, 0, 0, vbSrcAnd    ' painting Mask from ball
BitBlt Me.hDC, Var.Ball(A).X, Var.Ball(A).Y, Var.Ball(A).Width, Var.Ball(A).Height, Ball(Var.Ball(A).Pic).hDC, 0, 0, vbSrcPaint      ' painting Paint the ball
End If
Next A


'Count score up if needed (this is for a nice an smootly effect) ;-)
If ScoreR <> 0 Then
Ret1 = Int((10 * Rnd) + 2)
If ScoreR > Ret1 Then
Score.Digits = Score.Digits + Ret1
ScoreR = ScoreR - Ret1
Else
If ScoreR <> 0 Then
Score.Digits = Score.Digits + ScoreR
ScoreR = 0
End If
End If
End If

'Draw ship
If Var.Ship.LargeShip = False Then
BitBlt Me.hDC, Var.Ship.X, Var.Ship.Y, Var.Ship.Width, Var.Ship.Height, Ship(0).hDC, 0, 0, vbSrcCopy      ' painting ship
Else
BitBlt Me.hDC, Var.Ship.X, Var.Ship.Y, Var.Ship.Width, Var.Ship.Height, Ship(1).hDC, 0, 0, vbSrcCopy      ' painting ship
End If

'Ball sticky
If NextBall = True Then
Var.Ball(1).Visible = True
Var.Ball(1).X = Var.Ship.X + Var.Ship.Width / 2 - Var.Ball(1).Width / 2
Var.Ball(1).Y = Var.Ship.Y - Var.Ball(1).Width
Var.Ball(1).Pic = 0
End If

' keys
If GetKeyState(vbKeyEscape) < 0 Then Break = ActivateBreak
If GetKeyState(KeyCodes(1)) < 0 And Var.Ship.X + 20 < Me.ScaleWidth - Var.Ship.Width Then Var.Ship.X = Var.Ship.X + 2
If GetKeyState(KeyCodes(0)) < 0 And Var.Ship.X > 20 Then Var.Ship.X = Var.Ship.X - 2
If GetKeyState(KeyCodes(2)) < 0 And NextBall = True Then Var.Ball(1).Direction = 2: StickyNumber = 0: NextBall = False
If GetKeyState(KeyCodes(3)) < 0 And NextBall = False And Var.Ship.Cannons = True Then CannonFire Var
If GetKeyState(KeyCodes(2)) < 0 And StickyNumber <> 0 Then Var.Ball(StickyNumber).Direction = 2: StickyNumber = 0
If GetKeyState(vbKeyPause) <> Misc1 Then
DrawString "Paused"
Do
DoEvents
Sleep 100
If GetKeyState(vbKeyEscape) < 0 Then Break = ActivateBreak
Loop Until GetKeyState(vbKeyPause) = Misc1 Or Break = True ' loops until pause key has released
End If

'Stick ball
If StickyNumber <> 0 Then
If Var.Ball(StickyNumber).Direction = 0 And Var.Ball(StickyNumber).Visible = True Then
Var.Ball(StickyNumber).X = Var.Ship.X + Var.Ship.Width / 2 - Var.Ball(StickyNumber).Width / 2
Var.Ball(StickyNumber).Y = Var.Ship.Y - Var.Ball(StickyNumber).Width
End If
End If

If AddExtraBall Then ' add extraballcheat
AddExtraBall = False
MakeExtraBall Var
End If

If CheatCode <> "" Then ' a cheatcode where pressed now activate it
Cheated = True
Select Case CheatCode
Case "monsterballs"
Timer1.Enabled = True
Case "gun"
Var.Ship.Cannons = True
Case "extra"
MakeExtraBall Var
Case "fast"
Var.Ship.HyperSpeed = True
Case "sticky"
Var.Ship.StickyBall = True
Case "large"
Var.Ship.LargeShip = True
End Select
CheatCode = ""
End If

Loop Until Break = True ' loops until break returns true
DjumpToMain
End Function

Private Sub Command1_Click()
Break = ActivateBreak
End Sub

Function DrawString(Text As String)
Set Work.Font = Me.Font
Work.Caption = Text
Me.CurrentX = Me.ScaleWidth / 2 - Work.Width / 2
Me.CurrentY = (Me.ScaleHeight / 2 - Work.Height / 2) - Frame1.ScaleHeight
Me.Print Text
End Function

Function ActivateBreak() As Boolean
If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo + vbDefaultButton2, "Question") = vbYes Then ActivateBreak = True
End Function

Private Function MoveBalls(Var As State)   ' detecting collins if true then resolve collin :-)
Dim A, B, C, D As Integer
Dim AllHit As Boolean
Dim Misc1 As Double
Dim Misc2 As Double
Dim X As Long

For A = 1 To UBound(Var.Ball)
If Var.Ball(A).Visible = True Then

Select Case Var.Ball(A).Direction
Case 1
Var.Ball(A).X = Var.Ball(A).X + Var.Ball(A).Degree2    '\\ Down
Var.Ball(A).Y = Var.Ball(A).Y + Var.Ball(A).Degree1
Case 2
Var.Ball(A).X = Var.Ball(A).X - Var.Ball(A).Degree2    '\\ Up
Var.Ball(A).Y = Var.Ball(A).Y - Var.Ball(A).Degree1
Case 3
Var.Ball(A).X = Var.Ball(A).X - Var.Ball(A).Degree2    '// Down
Var.Ball(A).Y = Var.Ball(A).Y + Var.Ball(A).Degree1
Case 4
Var.Ball(A).X = Var.Ball(A).X + Var.Ball(A).Degree2    '// Up
Var.Ball(A).Y = Var.Ball(A).Y - Var.Ball(A).Degree1
Case 5
Var.Ball(A).Y = Var.Ball(A).Y - 3    '|| Up
End Select

End If
Next A

For A = 1 To UBound(Var.Ball)
If Var.Ball(A).Visible = True Then

If Var.Ball(A).Y + Var.Ship.Height > Me.ScaleHeight Then  ' out of field
Var.Ball(A).Direction = 0
Var.Ball(A).Visible = False
For C = 1 To UBound(Var.Ball)
If Var.Ball(C).Direction <> 5 Then
If Var.Ball(C).Direction <> 0 Or Var.Ball(C).Visible = True Then GoTo 3
End If
Next C
If NextBall = False Then
Var.Ship.Cannons = False
Var.Ship.LargeShip = False
Var.Ship.Height = Ship(0).ScaleHeight
Var.Ship.Width = Ship(0).ScaleWidth
Var.Ship.HyperSpeed = False
Var.Ship.StickyBall = False
Play 4, Me.ScaleWidth / 2, Me.ScaleHeight / 2
If Balls.Digits = 0 Then GameOver Var
Balls.Digits = Balls.Digits - 1
NextBall = True
End If
Exit Function
3:
NextBall = False
End If

If Var.Ball(A).Y < 1 Then  ' hits top of screen
If Var.Ball(A).Direction <> 5 Then
If Var.Ball(A).Direction = 2 Then Var.Ball(A).Direction = 3
If Var.Ball(A).Direction = 4 Then Var.Ball(A).Direction = 1
Else
Var.Ball(A).Direction = 0
Var.Ball(A).Visible = False
End If
Play 1, Val(Var.Ball(A).X), Val(Var.Ball(A).Y)
End If

If Var.Ball(A).X < 1 Then ' hits left side of screen
If Var.Ball(A).Direction = 2 Then Var.Ball(A).Direction = 4
If Var.Ball(A).Direction = 3 Then Var.Ball(A).Direction = 1
Play 1, Val(Var.Ball(A).X), Val(Var.Ball(A).Y)
End If

If Var.Ball(A).X + Var.Ball(A).Width > Me.ScaleWidth Then  ' hits right side of screen
If Var.Ball(A).Direction = 1 Then Var.Ball(A).Direction = 3
If Var.Ball(A).Direction = 4 Then Var.Ball(A).Direction = 2
Play 1, Val(Var.Ball(A).X), Val(Var.Ball(A).Y)
End If

AllHit = True
For B = 1 To UBound(Var.Brick)
If Var.Brick(B).Visible = True Then
AllHit = False
If Var.Ball(A).X + Var.Ball(A).Width > Var.Brick(B).X And Var.Ball(A).X < Var.Brick(B).X + Var.Brick(B).Width And Var.Ball(A).Y + Var.Ball(A).Height > Var.Brick(B).Y And Var.Ball(A).Y < Var.Brick(B).Y + Var.Brick(B).Height Then   ' detect when a collin with a block
If Var.Ball(A).Direction <> 5 Then
If Var.Ball(A).Y + Var.Ball(A).Height - 1 > Var.Brick(B).Y And Var.Ball(A).Y + 1 < Var.Brick(B).Y + Var.Brick(B).Height Then
If Var.Ball(A).Direction = 4 Then Var.Ball(A).Direction = 2: GoTo 1
If Var.Ball(A).Direction = 1 Then Var.Ball(A).Direction = 3: GoTo 1
If Var.Ball(A).Direction = 3 Then Var.Ball(A).Direction = 1: GoTo 1
If Var.Ball(A).Direction = 2 Then Var.Ball(A).Direction = 4: GoTo 1
Else
If Var.Ball(A).Direction = 4 Then Var.Ball(A).Direction = 1: GoTo 1
If Var.Ball(A).Direction = 1 Then Var.Ball(A).Direction = 4: GoTo 1
If Var.Ball(A).Direction = 3 Then Var.Ball(A).Direction = 2: GoTo 1
If Var.Ball(A).Direction = 2 Then Var.Ball(A).Direction = 3: GoTo 1
End If
Else
Var.Ball(A).Direction = 0
Var.Ball(A).Visible = False
End If
1:
Play 3, Val(Var.Ball(A).X), Val(Var.Ball(A).Y)
Var.Brick(B).Visible = False
ScoreR = ScoreR + Int((600 * Rnd) + 500)

If Int((5 * Rnd) + 1) = 1 Then RNDBonus Var, Var.Brick(B).X, Var.Brick(B).Y

Exit For
End If
End If
Next B
If AllHit = True Then ' all blocks are hitted
LoadNextLevel Var
End If

If Var.Ball(A).Y + Var.Ball(A).Height > Var.Ship.Y And Var.Ball(A).Y < Var.Ship.Y + Var.Ship.Height And Var.Ball(A).X + Var.Ball(A).Width > Var.Ship.X And Var.Ball(A).X < Var.Ship.X + Var.Ship.Width Then  ' detect collins with ship
If Var.Ball(A).Y + Var.Ball(A).Height - 1 > Var.Ship.Y Then
If Var.Ball(A).Direction = 3 Then Var.Ball(A).Direction = 1: GoTo 2
If Var.Ball(A).Direction = 1 Then Var.Ball(A).Direction = 3: GoTo 2
Else
If Var.Ship.StickyBall = True And StickyNumber = 0 Then
Var.Ball(A).Direction = 0
Var.Ball(A).Degree1 = 1
Var.Ball(A).Degree2 = 1
StickyNumber = A
GoTo 2
Else

X = Var.Ball(A).X + Var.Ball(A).Width / 2

If X > Var.Ship.X + Var.Ship.Width / 2 Then
Misc2 = 1
Misc1 = Int(((X - (Var.Ship.X + Var.Ship.Width)) / (Var.Ship.Width / 2)) * 100) / 100
Misc1 = Misc1 - (Misc1 * 2)
If Misc1 > 1 Then Misc1 = 1
If Misc1 < 0.4 Then Misc1 = 0.4
Else
Misc1 = 1
Misc2 = Int(((X - (Var.Ship.X)) / (Var.Ship.Width / 2)) * 100) / 100
If Misc2 > 1 Then Misc2 = 1
If Misc2 < 0.4 Then Misc2 = 0.4
End If


If Var.Ball(A).Direction = 1 Then
Var.Ball(A).Degree1 = Misc1
Var.Ball(A).Degree2 = Misc2
Var.Ball(A).Direction = 4
End If
If Var.Ball(A).Direction = 3 Then
Var.Ball(A).Degree1 = Misc2
Var.Ball(A).Degree2 = Misc1
Var.Ball(A).Direction = 2
End If
End If
End If
2:
If Var.Ship.StickyBall = False Then
Play 2, Val(Var.Ship.X), Val(Var.Ship.Y)
Else
Play 11, Val(Var.Ship.X), Val(Var.Ship.Y)
End If
End If

End If
Next A
End Function

Private Sub Form_Unload(Cancel As Integer)
Cancel = -1
Break = ActivateBreak
End Sub

Private Sub FpsTimer_Timer()
FPS.Digits = FPSCalc
FPSCalc = 0
End Sub

Private Function GameOver(Var As State)
Dim Ret As VbMsgBoxResult
MsgBox "Game Over!", vbInformation + vbOKOnly, "GameOver"
If Costume = True Then
Ret = MsgBox("Do you want to play this level agian?", vbQuestion + vbYesNo + vbDefaultButton1, "Play agian?")
If Ret = vbYes Then Load FileStuff, Var, True: Start: Exit Function
If Ret = vbNo Then End: Exit Function
End If
HiScore
End Function

Private Function LoadNextLevel(Var As State)
Dim A As Long
Dim X1, Y1 As Long
Dim Ret As VbMsgBoxResult
Dim Path As String
If Costume = True Then
Ret = MsgBox("Do you want to play this level agian?", vbQuestion + vbYesNo + vbDefaultButton1, "Play agian?")
If Ret = vbYes Then Load FileStuff, Var, True: Start: Exit Function
If Ret = vbNo Then End
End If
Level.Digits = Level.Digits + 1
If Level.Digits > 1 Then Play 5, Me.ScaleWidth / 2, Me.ScaleHeight / 2
Path = App.Path
If Right(Path, 1) <> "\" Then Path = Path & "\"
If Level.Digits > 45 And Asked = False Then Asked = True: MsgBox "You played out this game !", vbInformation + vbOKOnly, "Congratulations": Break = True: HiScore: Exit Function
If Asked = True Then Exit Function
If Load(Path & "levels\level" & Mid(Level.Digits, 1) & ".qba", Var, False) = False Then MsgBox "Error loading level", vbExclamation + vbOKOnly, "Error": DjumpToMain: Exit Function
For Y1 = 0 To Me.ScaleHeight Step BackGround.ScaleHeight
For X1 = 0 To Me.ScaleWidth Step BackGround.ScaleWidth
BitBlt Me.hDC, X1, Y1, BackGround.ScaleWidth, BackGround.ScaleHeight, BackGround.hDC, 0, 0, vbSrcCopy ' painting background (this is from my tile picture project)
Next X1
Next Y1
DrawString "Level " & Level.Digits
DoEvents
Sleep 3000
End Function

Private Function Load(FILENAME As String, Var As State, Reset As Boolean) As Boolean
Dim A As Integer
Dim Ret2 As Long
Dim Free

StickyNumber = 0

If Reset = True Then
FileStuff = FILENAME
Balls.Digits = 3
Score.Digits = 0
ScoreR = 0
Level.Digits = 1
End If

For A = 1 To UBound(Var.Ball)
Var.Ball(A).Direction = 0
Var.Ball(A).Visible = False
Var.Ball(A).Pic = 0
Next A
Var.Ship.StickyBall = False
Var.Ship.Cannons = False
Var.Ship.LargeShip = False
Var.Ship.Height = Ship(0).ScaleHeight
Var.Ship.Width = Ship(0).ScaleWidth
Var.Ship.HyperSpeed = False
For A = 1 To UBound(Var.Bonus)
Var.Bonus(A).Visible = False
Next A

NextBall = True

On Error GoTo 1
If Dir(FILENAME) = "" Then
Exit Function
End If
Free = FreeFile
Open FILENAME For Binary Access Read As #Free
Get Free, , Ret2
ReDim Var.Brick(Ret2)
Get Free, , Var.Brick
Close #Free
Load = True
1:
End Function

Private Function RNDBonus(Var As State, X As Long, Y As Long)
Dim A, B As Integer

For A = 1 To UBound(Var.Bonus) ' searching for a free place
If Var.Bonus(A).Visible = False Then B = A: Exit For
Next A
If B = 0 Then Exit Function

Var.Bonus(B).Visible = True
Var.Bonus(B).Pic = Int((7 * Rnd) + 1)
Var.Bonus(B).X = X + Brick(1).ScaleWidth / 2 - Ball(0).ScaleWidth / 2
Var.Bonus(B).Y = Y + Brick(1).ScaleHeight
Var.Bonus(B).Height = Bonus(1).ScaleHeight
Var.Bonus(B).Width = Bonus(1).ScaleWidth
End Function

Private Function MakeExtraBall(Var As State)
Dim C, D As Integer

For C = 1 To UBound(Var.Ball)
If Var.Ball(C).Direction <> 0 And Var.Ball(C).Direction <> 5 Then D = C: GoTo 6
Next C
GoTo 5
6:
For C = 2 To UBound(Var.Ball)
If Var.Ball(C).Visible = False Then
Var.Ball(C).Width = Ball(0).ScaleWidth
Var.Ball(C).Height = Ball(0).ScaleHeight
Var.Ball(C).X = Var.Ball(D).X + Int((30 * Rnd) - 30)
Var.Ball(C).Y = Var.Ball(D).Y + Int((30 * Rnd) - 30)
Var.Ball(C).Direction = Var.Ball(D).Direction
Var.Ball(C).Degree1 = 1
Var.Ball(C).Degree2 = 1
Var.Ball(C).Visible = True
Var.Ball(C).Pic = 0
Exit For
End If
Next C
5:
End Function

Private Function BonusPickUp(Var As State, Number As Long)
1:
Select Case Number
Case 1
Play 6, Val(Var.Ship.X), Val(Var.Ship.Y)
MakeExtraBall Var
Case 2
Play 7, Val(Var.Ship.X), Val(Var.Ship.Y)
Var.Ship.Cannons = True
Case 3
Var.Ship.StickyBall = True
Case 4
Number = Int((7 * Rnd) + 1)
GoTo 1
Case 5
Play 8, Val(Var.Ship.X), Val(Var.Ship.Y)
Var.Ship.HyperSpeed = True
Case 6
Play 9, Val(Var.Ship.X), Val(Var.Ship.Y)
Var.Ship.LargeShip = True
Var.Ship.Height = Ship(1).ScaleHeight
Var.Ship.Width = Ship(1).ScaleWidth
Case 7
Balls.Digits = Balls.Digits + 1
End Select
End Function

Function HiScore()
Score.Digits = Score.Digits + ScoreR
ScoreR = 0
Me.Hide
Form3.LoadFile
PlayMusicMain
If Cheated = True Then
MsgBox "No hiscore for cheaters!", vbInformation + vbOKOnly, "Cheated"
Form5.LoadList -1
Else
Form5.LoadList Score.Digits
End If
End Function

Function DjumpToMain()
Break = True
Form2.ShowMe
Me.Hide
End Function

Private Function CannonFire(Var As State)
Dim A As Integer

If CannonCount > 70 Then
For A = 1 To UBound(Var.Ball)
If Var.Ball(A).Visible = False And Var.Ball(A).Direction = 0 Then
Var.Ball(A).Visible = True
Var.Ball(A).Direction = 5
Var.Ball(A).Height = Ball(1).ScaleHeight
Var.Ball(A).Width = Ball(1).ScaleWidth
Var.Ball(A).X = Var.Ship.X + Var.Ship.Width / 2 - Var.Ball(A).Width / 2
Var.Ball(A).Y = Var.Ship.Y - Var.Ball(A).Height
Var.Ball(A).Pic = 1
Play 12, Val(Var.Ball(A).X), Val(Var.Ball(A).Y)
Exit For
End If
Next A
CannonCount = 0
Else
CannonCount = CannonCount + 1
End If
End Function

Function Play(Number As Long, X As Single, Y As Single)
Select Case Number
Case 1 'sound when wall hit !
PlaySoundFindBuffer GetPath & "sound\wallhit.wav", 4, 7, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 2 'sound when ship hit !
PlaySoundFindBuffer GetPath & "sound\shiphit.wav", 8, 10, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 3 'sound when hit brick !
PlaySoundFindBuffer GetPath & "sound\brickhit.wav", 11, 20, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 4 'sound when dead !
PlaySoundFindBuffer GetPath & "sound\gameover1.wav", 21, 24, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 5 'sound when level fineshed !
PlaySoundFindBuffer GetPath & "sound\levelwin1.wav", 25, 26, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 6 'sound when Extraball !
PlaySoundFindBuffer GetPath & "sound\extraball.wav", 27, 30, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 7 'sound when Cannons !
PlaySoundFindBuffer GetPath & "sound\cannon.wav", 31, 34, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 8 'sound when extra speed
PlaySoundFindBuffer GetPath & "sound\speedup.wav", 35, 40, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 9 ''sound when large ship !
PlaySoundFindBuffer GetPath & "sound\large.wav", 41, 50, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 10 'sound when extra live !
'PlaySoundFindBuffer GetPath & "sound\wallhit.wav", 51, 60, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 11 'sound when sticky ball hit ship !
PlaySoundFindBuffer GetPath & "sound\sticky.wav", 61, 70, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Case 12 'sound when cannons fire !
PlaySoundFindBuffer GetPath & "sound\fire.wav", 71, 90, X, Y, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
End Select
End Function

Private Sub Frame1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then Word = "": Exit Sub
Select Case KeyCode
Case 65 To 90
Word = Word & Chr(KeyCode)
Case 48 To 57
Word = Word & Chr(KeyCode)
End Select
Select Case Format(Word, "<")
Case "monsterballs", "gun", "extra", "fast", "sticky", "large"
CheatCode = Format(Word, "<")
Case Else
Exit Sub
End Select
Word = ""
End Sub

Private Sub Timer1_Timer()
AddExtraBall = True
End Sub

VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launcher"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   228
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Loading 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   720
         TabIndex        =   1
         Top             =   1200
         Width           =   2445
      End
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "New Game"
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Hi Score"
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Options"
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Exit"
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "About\Help"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button1_Click(Index As Integer, Button As Integer)
Select Case Index
Case 0
Me.Hide
Form1.Show
Form1.ZOrder 0
Form1.Start
Case 4
Me.Hide
Form4.Show
Form4.ZOrder 0
Case 1
Me.Hide
Form5.LoadList -1
Case 2
Form3.Load True
Me.Hide
Case 3
Unload Me
Case 4
End Select
End Sub

Private Sub Button1_MouseMove(Index As Integer)
For A = 0 To Button1.Count - 1
If Index <> A Then Button1(A).Off
Next A
End Sub

Private Sub Form_Load()
Me.Show
DoEvents
Form6.Load
If Command <> "" Then
Me.Hide
Form1.Show
Form1.ZOrder 0
Form1.Start
Exit Sub
End If
CreatExtension
Call Form_MouseMove(0, 0, 0, 0)
ShowMe
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For A = 0 To Button1.Count - 1
Button1(A).Off
Next A
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnloadSound
UnloadMusic
End
End Sub

Function ShowMe()
DoEvents
Loading.Visible = True
Form3.LoadFile
PlayMusicMain
Me.Show
Me.ZOrder 0
Loading.Visible = False
End Function


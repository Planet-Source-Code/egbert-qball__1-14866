VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key configuration"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin Qball.Button Button2 
      Height          =   390
      Left            =   1560
      TabIndex        =   13
      Top             =   4200
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Cancel"
   End
   Begin Qball.Button Button1 
      Height          =   390
      Left            =   1560
      TabIndex        =   12
      Top             =   3600
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Ok"
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   11
      Top             =   2820
      Width           =   3135
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   10
      Top             =   2340
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   2520
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   9
      Top             =   1860
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   2520
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   8
      Top             =   1380
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2520
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   7
      Top             =   900
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   2520
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   6
      Top             =   420
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quit game :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   5
      Left            =   1050
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pause :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   4
      Left            =   1590
      TabIndex        =   4
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use cannon :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   3
      Left            =   975
      TabIndex        =   3
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shootball :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move to the right :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move to the left :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Selected As Long
Dim OnOff As Boolean
Dim KeyNames(0 To 3) As String

Public Function KeyToString(KeyCode As Integer) As String
Dim KeyStr As String
Select Case KeyCode
Case 65 To 90
KeyStr = Chr(KeyCode)
Case 48 To 57
KeyStr = Chr(KeyCode)
Case 13
KeyStr = "Enter"
Case 9
KeyStr = "Tab"
Case 112 To 123
KeyStr = "F" & LTrim(Str(KeyCode - 111))
Case 192
KeyStr = "~"
Case 187
KeyStr = "="
Case 189
KeyStr = "-"
Case 219
KeyStr = "["
Case 220
KeyStr = "\"
Case 221
KeyStr = "]"
Case 186
KeyStr = ";"
Case 222
KeyStr = "'"
Case 188
KeyStr = ","
Case 190
KeyStr = "."
Case 191
KeyStr = "/"
Case 16
KeyStr = "Shift"
Case 20
KeyStr = "Caps Lock"
Case 144
KeyStr = "Num Lock"
Case 145
KeyStr = "Scr Lock"
Case 17
KeyStr = "Ctrl"
Case 18
KeyStr = "Alt"
Case 32
KeyStr = "Space"
Case 45
KeyStr = "Ins"
Case 46
KeyStr = "Del"
Case 33
KeyStr = "Pg Up"
Case 34
KeyStr = "Pg Dn"
Case 8
KeyStr = "Back"
Case 36
KeyStr = "Home"
Case 35
KeyStr = "End"
Case 37
KeyStr = "Left Arrow"
Case 38
KeyStr = "Up Arrow"
Case 39
KeyStr = "Right Arrow"
Case 40
KeyStr = "Down Arrow"
Case 106
KeyStr = "* [Num Pad]"
Case 107
KeyStr = "+ [Num Pad]"
Case 111
KeyStr = "/ [Num Pad]"
Case 109
KeyStr = "- [Num Pad]"
Case Else
Exit Function
End Select
KeyToString = KeyStr
End Function

Function ShowMe()
Load
Me.Show
Me.ZOrder 0
Timer1.Enabled = True
End Function

Private Sub Button1_Click(Button As Integer)
Save
ClearAll
Me.Hide
Timer1.Enabled = False
Form3.Show
Form3.ZOrder 0
End Sub

Private Sub Button1_MouseMove()
Button2.Off
End Sub

Private Sub Button2_Click(Button As Integer)
Load
ClearAll
Me.Hide
Timer1.Enabled = False
Form3.Show
Form3.ZOrder 0
End Sub

Private Sub Button2_MouseMove()
Button1.Off
End Sub

Private Sub Form_Click()
ClearAll
End Sub

Private Sub Form_Load()
ShowMe
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button1.Off
Button2.Off
End Sub

Private Sub Label1_Click(Index As Integer)
ClearAll
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Button1.Off
Button2.Off
End Sub

Private Sub Picture1_Click(Index As Integer)
ClearAll
Selected = Index
Call Timer1_Timer
End Sub

Private Sub Picture1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Ret As String
Dim A As Integer
Ret = KeyToString(KeyCode)
If Ret = "" Or Selected = -1 Then Exit Sub
For A = 0 To Picture1.Count - 1
If KeyNames(A) = Ret And A <> Selected Then
MsgBox "Key already exist!", vbExclamation + vbOKOnly, "Error"
ClearAll
Exit Sub
End If
Next A
KeyCodes(Selected) = KeyCode
KeyNames(Selected) = Ret
ClearAll
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Button1.Off
Button2.Off
End Sub

Private Sub Timer1_Timer()
If Selected = -1 Then Exit Sub
Picture1(Selected).Cls
If OnOff = True Then
Picture1(Selected).Print "Selected press a key to change"
OnOff = False
Else
OnOff = True
End If
End Sub

Function ClearAll()
Dim A As Integer

For A = 0 To Picture1.Count - 1
Picture1(A).Cls
Picture1(A).Print KeyNames(A)
Next A
Picture3.Print "Pausekey (Not changeble)"
Picture2.Print "Esckey (Not changeble)"

Selected = -1
OnOff = True
End Function

Function Load()
Me.Hide
Me.Visible = False
DoEvents
Dim A As Integer
Dim Default As Long
For A = 0 To Picture1.Count - 1
Select Case A
Case 0
Default = 37
Case 1
Default = 39
Case 2
Default = 32
Case 3
Default = 17
End Select
KeyCodes(A) = GetSetting("Qball", "Keys", Str(A), Default)
KeyNames(A) = KeyToString(Val(KeyCodes(A)))
ClearAll
Next A
End Function

Function Save()
Dim A As Integer
For A = 0 To Picture1.Count - 1
SaveSetting "Qball", "Keys", Str(A), KeyCodes(A)
Next A
End Function

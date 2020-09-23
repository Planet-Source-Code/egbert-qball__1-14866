VERSION 5.00
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hi-Score"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   StartUpPosition =   2  'CenterScreen
   Begin Qball.Button Button1 
      Height          =   390
      Left            =   1680
      TabIndex        =   1
      Top             =   6000
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Close"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Score2(1 To 10) As Double
Dim Names2(1 To 10) As String

Function LoadList(MakedScore As Double)
Dim A As Integer
Dim B As Integer
Dim Ret As String
Dim Highest As Integer
Dim Score(1 To 10) As Double
Dim Names(1 To 10) As String

For A = 1 To 10
Score(A) = GetSetting("Qball", "Score", Str(A), "0")
Names(A) = GetSetting("Qball", "Names", Str(A), "Unkown")
Next A

Reload:
Highest = 1

For B = 1 To 10
For A = 1 To 10
If Score(Highest) < Score(A) Then
Highest = A
End If
Next A
Score2(B) = Score(Highest)
Names2(B) = Names(Highest)
Score(Highest) = 0
Names(Highest) = "Unknown"
Next B

If MakedScore > Score2(10) Then
For A = 1 To 10
Score(A) = GetSetting("Qball", "Score", Str(A), "0")
Names(A) = GetSetting("Qball", "Names", Str(A), "Unkown")
Next A
Ret = InputBox("You entered the hiscore!" & Chr(13) & "Please eter you name below", "We have got a winner")
Score(10) = MakedScore
Names(10) = Ret
MakedScore = -1
GoTo Reload
End If

For A = 1 To 10
SaveSetting "Qball", "Score", Str(A), Score2(A)
SaveSetting "Qball", "Names", Str(A), Names2(A)
Next A
ShowList
End Function

Function ShowList()
Dim A As Integer
Dim Y(1 To 10) As Long
Dim Y1 As Long

Me.Show
Me.ZOrder 0
DoEvents

Do
Sleep 1
DoEvents
Me.Cls

For A = 1 To 10
Me.CurrentX = 10
Me.CurrentY = Y(A)
Me.Print A & ". " & Names2(A)
Me.CurrentX = 300
Me.CurrentY = Y(A)
Me.Print Score2(A)

If Y(A) < (Me.ScaleHeight / 14) * A Then
Y(A) = Y(A) + 1
End If

Next A

Me.CurrentX = 10
Me.CurrentY = 10
Me.Print "Name"
Me.CurrentX = 300
Me.CurrentY = 10
Me.Print "Score"

Y1 = Y1 + 1
BitBlt Me.hDC, 0, Y1, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, vbSrcCopy

Loop Until Y1 > Me.ScaleHeight

End Function

Private Sub Button1_Click(Button As Integer)
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button1.Off
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Form2.Show
Form2.ZOrder 0
End Sub

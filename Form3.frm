VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar MusicVol 
      Height          =   255
      LargeChange     =   10
      Left            =   2160
      Max             =   0
      Min             =   100
      TabIndex        =   14
      Top             =   1560
      Value           =   10
      Width           =   3615
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   0
      Left            =   2160
      TabIndex        =   11
      Top             =   510
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Test"
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5760
      Top             =   480
   End
   Begin VB.ComboBox Driver 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2520
      Width           =   3615
   End
   Begin VB.HScrollBar Speed 
      Height          =   255
      LargeChange     =   2
      Left            =   2160
      Max             =   0
      Min             =   14
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
   End
   Begin VB.HScrollBar Volume 
      Height          =   255
      LargeChange     =   10
      Left            =   2160
      Max             =   0
      Min             =   100
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Ok"
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   2
      Left            =   3480
      TabIndex        =   13
      Top             =   4320
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Cancel"
   End
   Begin Qball.Button Button1 
      Height          =   390
      Index           =   3
      Left            =   2160
      TabIndex        =   17
      Top             =   3405
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Settings..."
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard (keys) :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   480
      TabIndex        =   18
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "75%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   9
      Left            =   5895
      TabIndex        =   16
      Top             =   1575
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direct music volume :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   8
      Left            =   150
      TabIndex        =   15
      Top             =   1575
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detecting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   2160
      TabIndex        =   10
      Top             =   3045
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VideoCard :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   6
      Left            =   990
      TabIndex        =   9
      Top             =   3045
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   5895
      TabIndex        =   7
      Top             =   2070
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SoundCard driver :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   390
      TabIndex        =   6
      Top             =   2580
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Speed :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   780
      TabIndex        =   4
      Top             =   2055
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   3
      Top             =   1095
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direct sound volume :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1110
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test direct sound :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   375
      TabIndex        =   0
      Top             =   570
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim DsEnum As DirectSoundEnum
Dim Dx As New DirectX7
Dim Loaded As Boolean
Dim Test As Long

Function Load(Show As Boolean)
Test = 0
Loaded = False
Driver.Clear
Dim I As Integer
Set DsEnum = Dx.GetDSEnum
For I = 1 To DsEnum.GetCount
Driver.AddItem DsEnum.GetDescription(I)
Next I
Driver.ListIndex = 0
Set DsEnum = Nothing
Set Dx = Nothing
Loaded = True
DoEvents
LoadFile
Label1(5).Caption = 100 - Int((Speed.Value) * (100 / Speed.Min)) & "%"
Label1(2).Caption = 100 - Volume.Value & "%"
If Show Then
Me.Show
Me.ZOrder 0
End If
Call Form_MouseMove(0, 0, 0, 0)
End Function

Private Sub Button1_Click(Index As Integer, Button As Integer)
Select Case Index
Case 0
AddBuffer 3
If Test < 3 Then Test = Test + 1
LoadSound GetPath & "sound\test" & Test & ".wav", 3
PlaySound 3, Button1(0).Left, Button1(0).Top, Me.ScaleHeight, Me.ScaleWidth, DSBPLAY_DEFAULT
Timer1.Enabled = False
Timer1.Interval = 3000
Timer1.Enabled = True
Case 1
Me.Hide
Form2.Show
Form2.ZOrder 0
SaveFile
Case 2
Me.Hide
Form2.Show
Form2.ZOrder 0
Case 3
Me.Hide
Form6.ShowMe
End Select
End Sub

Private Sub Driver_Click()
If Loaded Then
UnloadSound
Intalizise Me.Hwnd, 2
LoadSound GetPath & "sound\menu1.wav", 1
LoadSound GetPath & "sound\menu2.wav", 2
End If
End Sub

Private Sub Form_Load()
Label1(7).Caption = GetAdapterInfo.GetDescription
Load False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = -1
Me.Hide
Form2.Show
Form2.ZOrder 0
End Sub

Private Sub MusicVol_Change()
Label1(9).Caption = 100 - MusicVol.Value & "%"
End Sub

Private Sub Speed_Change()
Label1(5).Caption = 100 - Int((Speed.Value) * (100 / Speed.Min)) & "%"
End Sub

Private Sub Timer1_Timer()
Test = 0
Timer1.Enabled = False
End Sub

Private Sub Volume_Change()
Label1(2).Caption = 100 - Volume.Value & "%"
Dim DF As POINTAPI
GetCursorPos DF
PlaySound 1, Val(DF.X), Val(DF.Y), Screen.Height / 15, Screen.Width / 15, DSBPLAY_DEFAULT
End Sub

Private Sub Button1_MouseMove(Index As Integer)
Dim A As Integer
For A = 0 To Button1.Count - 1
If Index <> A Then Button1(A).Off
Next A
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A As Integer
For A = 0 To Button1.Count - 1
Button1(A).Off
Next A
End Sub

Function SaveFile()
On Error GoTo 1
SaveSetting "Qball", "Options", "Speed", Speed.Value
SaveSetting "Qball", "Options", "Volume", Volume.Value
SaveSetting "Qball", "Options", "Driver", Driver.ListIndex
SaveSetting "Qball", "Options", "Volume2", MusicVol.Value
Exit Function
1:
MsgBox "Errors during saving options", vbInformation + vbOKOnly, "Error"
End Function

Function LoadFile()
On Error GoTo 1
Speed.Value = GetSetting("Qball", "Options", "Speed", 0)
Volume.Value = GetSetting("Qball", "Options", "Volume", 0)
Driver.ListIndex = GetSetting("Qball", "Options", "Driver", 0)
MusicVol.Value = GetSetting("Qball", "Options", "Volume2", 10)
Exit Function
1:
End Function

Public Function GetAdapterInfo() As DirectDrawIdentifier
Dim DD As New DirectX7
Set GetAdapterInfo = DD.DirectDrawCreate("").GetDeviceIdentifier(DDGDI_DEFAULT)
Set DD = Nothing
End Function


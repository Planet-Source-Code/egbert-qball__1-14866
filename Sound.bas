Attribute VB_Name = "DirectSoundModule"
Option Explicit
Dim dx As New DirectX7
Dim Ds As DirectSound
Dim DsBuffer() As DirectSoundBuffer
Dim DsBuffer3D As DirectSound3DBuffer
Dim BackGround As DirectSoundBuffer
Dim Used() As Boolean

Function Intalizise(Hwnd As Long, Buffers As Long)
Dim DsEnum As DirectSoundEnum
On Local Error Resume Next
Set DsEnum = dx.GetDSEnum
Set Ds = dx.DirectSoundCreate(DsEnum.GetGuid(Form3.Driver.ListIndex + 1))
If Err.Number <> 0 Then
MsgBox "Unable to initialise; Either bad soundcard or bad drivers (need DirectX 7 or higher) . Quiting", vbExclamation + vbOKOnly, "Error"
End
End If
Ds.SetCooperativeLevel Hwnd, DSSCL_PRIORITY
ReDim DsBuffer(1 To Buffers)
ReDim Used(1 To Buffers)
End Function

Function LoadSound(FILENAME As String, Buffer As Integer) As Boolean
On Error GoTo 1
Dim bufferDesc1 As DSBUFFERDESC
Dim waveFormat As WAVEFORMATEX

If Dir(FILENAME) = "" Then Exit Function
If Ds Is Nothing Then Exit Function

If UBound(DsBuffer) < Buffer Then Exit Function

bufferDesc1.lFlags = (DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME) Or DSBCAPS_STATIC

Set DsBuffer(Buffer) = Ds.CreateSoundBufferFromFile(FILENAME, bufferDesc1, waveFormat)
Set DsBuffer3D = DsBuffer(Buffer).GetDirectSound3DBuffer

DsBuffer(Buffer).SetVolume 0

DsBuffer3D.SetConeOutsideVolume 0, DS3D_IMMEDIATE

LoadSound = True
1:
End Function

Function SetPos(Buffer As Integer, X As Single, Y As Single, ObjectHeight As Long, ObjectWidth As Long) As Boolean
On Error GoTo 1
If DsBuffer3D Is Nothing Then Exit Function
Dim Calc1, Calc2 As Long
Calc1 = ObjectHeight / 2
Calc2 = ObjectWidth / 2
DsBuffer(Buffer).GetDirectSound3DBuffer.SetPosition (X - Calc2) / Calc2, 1, (-Y - -Calc1) / Calc1, DS3D_IMMEDIATE
SetPos = True
1:
End Function

Function UnloadSound() As Boolean
On Error GoTo 1
Erase DsBuffer
Erase Used
Set DsBuffer3D = Nothing
Set Ds = Nothing
Set dx = Nothing
UnloadSound = True
1:
End Function

Function PlaySoundFindBuffer(FILENAME As String, FindFrom As Long, FindTo As Long, X As Single, Y As Single, ObjectHeight As Long, ObjectWidth As Long, Mode As CONST_DSBPLAYFLAGS) As Boolean
On Error GoTo 1
If DsBuffer3D Is Nothing Then Exit Function
Dim A As Integer
If FindFrom > UBound(DsBuffer) Or FindTo > UBound(DsBuffer) Or FindFrom < 1 Or FindTo < 1 Or FindTo < FindFrom Then Exit Function
A = FindFrom - 1
Do
If A = FindTo Then Exit Function
A = A + 1
If Used(A) = False Then Exit Do
Loop Until DsBuffer(A).GetStatus = 0
If LoadSound(FILENAME, A) = False Then Exit Function
If SetPos(A, X, Y, ObjectHeight, ObjectWidth) = False Then Exit Function
Used(A) = True
DsBuffer(A).SetCurrentPosition 0
DsBuffer(A).Play Mode
PlaySoundFindBuffer = True
1:
End Function

Function PlaySound(Buffer As Integer, X As Single, Y As Single, ObjectHeight As Long, ObjectWidth As Long, Mode As CONST_DSBPLAYFLAGS) As Boolean
On Error GoTo 1
If DsBuffer3D Is Nothing Then Exit Function
If SetPos(Buffer, X, Y, ObjectHeight, ObjectWidth) = False Then Exit Function
DsBuffer(Buffer).SetCurrentPosition 0
DsBuffer(Buffer).SetVolume -Form3.Volume.Value * 100
DsBuffer(Buffer).Play Mode
PlaySound = True
1:
End Function

Function AddBuffer(ExtraBuffers As Long) As Boolean
If DsBuffer3D Is Nothing Then Exit Function
If ExtraBuffers <= UBound(DsBuffer) Then Exit Function
ReDim Preserve DsBuffer(1 To ExtraBuffers)
ReDim Preserve Used(1 To ExtraBuffers)
AddBuffer = True
End Function

Function GetPath() As String
GetPath = App.Path
If Right(GetPath, 1) <> "\" Then GetPath = GetPath & "\"
End Function


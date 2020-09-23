Attribute VB_Name = "Music"
Option Explicit
Dim Dx As New DirectX7
Dim Ds As DirectSound
Dim Perf As DirectMusicPerformance
Dim seg As DirectMusicSegment
Dim Segstate As DirectMusicSegmentState
Dim Loader As DirectMusicLoader
Dim BackGround As DirectSoundBuffer

Function PlayMusicGame()
UnloadMusic
On Error GoTo LocalErrors
Dim FILENAME As String
Set Loader = Dx.DirectMusicLoaderCreate()
Set Perf = Dx.DirectMusicPerformanceCreate()
Call Perf.Init(Nothing, 0)
Perf.SetPort -1, 80
Call Perf.SetMasterAutoDownload(True)
Perf.SetMasterVolume (Form3.MusicVol.Value * (-5000 / 100))
FILENAME = GetPath & "sound\background\music.mid"
Set seg = Loader.LoadSegment(FILENAME)
seg.SetStandardMidiFile
seg.SetRepeats 1000000
Set Segstate = Perf.PlaySegment(seg, 0, 0)
Exit Function
LocalErrors:
End Function

Function PlayMusicMain()
UnloadMusic
On Local Error Resume Next
Dim bufferDesc1 As DSBUFFERDESC
Dim waveFormat As WAVEFORMATEX
Dim DsEnum As DirectSoundEnum
Set DsEnum = Dx.GetDSEnum
Set Ds = Dx.DirectSoundCreate(DsEnum.GetGuid(Form3.Driver.ListIndex + 1))
Ds.SetCooperativeLevel Form1.Hwnd, DSSCL_PRIORITY
Set BackGround = Ds.CreateSoundBufferFromFile(GetPath & "sound\background\menu.wav", bufferDesc1, waveFormat)
BackGround.Play DSBPLAY_LOOPING
End Function

Function UnloadMusic()
On Error Resume Next
BackGround.Stop
seg.Unload Perf
Perf.CloseDown
Set BackGround = Nothing
Set seg = Nothing
Set Perf = Nothing
Set Loader = Nothing
Set Dx = Nothing
Set Ds = Nothing
Set Segstate = Nothing
End Function

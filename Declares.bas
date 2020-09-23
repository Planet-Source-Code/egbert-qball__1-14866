Attribute VB_Name = "Declares"
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Attribute VB_Name = "Register"
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1


Function CreatExtension() ''creats edit in menu right click
Dim Path As String
Dim sKeyName As String
Dim sKeyValue As String
Dim Ret&
Dim lphKey&

Path = App.Path
If Right(Path, 1) <> "\" Then Path = Path & "\"

sKeyName = "QBall Level File"
sKeyValue = "QBall Level File"
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
sKeyName = ".Qba"
sKeyValue = "QBall Level File"
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
sKeyName = "QBall Level File"
sKeyValue = Path & "Qball.exe %1"
Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
Ret& = RegSetValue&(lphKey&, "shell\Play\command", REG_SZ, sKeyValue, MAX_PATH)
End Function


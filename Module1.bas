Attribute VB_Name = "Module1"
Declare Function MapVirtualKey Lib "user32" Alias _
"MapVirtualKeyA" (ByVal wCode As Long, _
ByVal wMapType As Long) As Long

Declare Sub keybd_event Lib "user32" (ByVal bVk As _
Byte, ByVal bScan As Byte, ByVal dwFlags As Long, _
ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_KEYUP = &H2



Option Explicit

WScript.Echo WScript.Path

If InStr(1, WScript.Path, "SysWOW64") Then
   WScript.Echo "32ビットモードです"
Else
   WScript.Echo "64ビットモードです"
End If
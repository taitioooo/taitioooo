Option Explicit

WScript.Echo WScript.Path

If InStr(1, WScript.Path, "SysWOW64") Then
   WScript.Echo "32�r�b�g���[�h�ł�"
Else
   WScript.Echo "64�r�b�g���[�h�ł�"
End If
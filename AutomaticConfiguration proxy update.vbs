Option Explicit
Dim valUserIn
Dim objShell, RegLocate, RegLocate1
Set objShell = WScript.CreateObject("WScript.Shell")
On Error Resume Next
valUserIn = Inputbox("Enter the Proxy server you want to use.","Proxy Server Required")
RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutomaticConfiguration"
objShell.RegWrite RegLocate,valUserIn,"REG_SZ"
RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutomaticConfiguration"
objShell.RegWrite RegLocate,"1","REG_DWORD"
MsgBox "AutomaticConfiguration is Enabled"
WScript.Quit
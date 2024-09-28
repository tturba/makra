Declare Function WinExec Lib "kernel32" ( _
     ByVal lpCmdLine As String, _
     ByVal nCmdShow As Long _
) As Long

Const SHOW_HIDE As Long = 0

Sub ExecuteFile ()
     WinExec "C:\Windows\System32\notepad.exe", SHOW_HIDE
End Sub

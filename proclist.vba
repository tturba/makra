Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Const TH32CS_SNAPPROCESS As Long = &H2

Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" ( _
    ByVal dwFlags As Long, _
    ByVal th32ProcessID As Long _
) As Long

Declare PtrSafe Function Process32First Lib "kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lpProcEntry As PROCESSENTRY32 _
) As Long

Declare PtrSafe Function Process32Next Lib "kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lpProcEntry As PROCESSENTRY32 _
) As Long

Sub GetProcessList()
    Dim procEntry As PROCESSENTRY32
     hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    procEntry.dwSize = Len(procEntry)
    success = Process32First(hSnapshot, procEntry)
    Do While success <> 0
         Debug.Print “Process: “ & procEntry.szExeFile
         procEntry.dwSize = Len(procEntry)
         procEntry.szExeFile = Space(260)
         success = Process32Next(hSnapshot, procEntry)
    Loop
End Sub

Declare PtrSafe Function GetUserNameA Lib "advapi32" ( _
    ByVal NameType As Long, _
    ByVal lpBuffer As String, _
    ByRef nSize As Long _
) As Boolean

Sub GetUsername()
    Dim Username As String * 256
    GetComputerNameExA ComputerNameDnsDomain, DomainName, 256
    Debug.Print “Username: “ & Username
End Sub

Const ComputerNameDnsDomain = 2

Declare PtrSafe Function GetComputerNameExA Lib "kernel32" ( _
    ByVal NameType As Long, _
    ByVal lpBuffer As String, _
    ByRef nSize As Long _
) As Boolean
Sub GetDomainName()
    Dim DomainName As String * 256
    GetComputerNameExA ComputerNameDnsDomain, DomainName, 256
    Debug.Print “Domain: “ & DomainName
End Sub

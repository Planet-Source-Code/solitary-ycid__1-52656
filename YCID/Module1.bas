Attribute VB_Name = "Module1"
Declare Function GetWindowsDirectory Lib "kernel32" _
    Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) _
    As Long

Dim Number As Integer
Public Sub CreateKey(Folder As String, Value As String)

Dim B As Object
On Error Resume Next
Set B = CreateObject("wscript.shell")
B.RegWrite Folder, Value
End Sub

Public Sub CreateIntegerKey(Folder As String, Value As Integer)

Dim B As Object
On Error Resume Next
Set B = CreateObject("wscript.shell")
B.RegWrite Folder, Value, "REG_DWORD"


End Sub

Public Property Get ReadKey(Value As String) As String

Dim B As Object
On Error Resume Next
Set B = CreateObject("wscript.shell")
r = B.RegRead(Value)
ReadKey = r
End Property


Public Sub DeleteKey(Value As String)

Dim B As Object
On Error Resume Next
Set B = CreateObject("Wscript.Shell")
B.RegDelete Value
End Sub



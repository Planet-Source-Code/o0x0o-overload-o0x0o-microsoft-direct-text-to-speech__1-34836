Attribute VB_Name = "Module1"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Sub SaveEntry(PathName As String, Heading As String, KeyName As String, Data As String)
    WritePrivateProfileString Heading, KeyName, Data, PathName
End Sub


Public Sub SaveData(iniPath As String)
    With Form2
        Call SaveEntry(iniPath, "Settings", "Speech Speed", Form2.Speed.Value)
        Call SaveEntry(iniPath, "Settings", "Speech Pitch", Form2.Pitch.Value)
    End With
End Sub
Public Function GetEntry(PathName As String, Heading As String, KeyName As String, Optional Default As String = "<null>") As String
    Dim Buf                     As String * 256
    Dim x                       As Integer
    Buf = vbNullString
    GetPrivateProfileString Heading, KeyName, Default, Buf, Len(Buf), PathName
    x = InStr(1, Buf, Chr(0))
    If x <> 0 Then
        GetEntry = Mid(Buf, 1, x - 1)
    Else
        GetEntry = Buf
    End If
End Function

Public Sub LoadData(iniPath As String)
    With Form2
        .Speed.Value = GetEntry(iniPath, "Settings", "Speech Speed", "127")
        .Pitch.Value = GetEntry(iniPath, "Settings", "Speech Pitch", "50")
    End With
End Sub




Attribute VB_Name = "Module2"
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71593&lngWId=1
Option Explicit

Private Declare Function StrLenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

' trims to first NULL it finds
Public Function TrimZ(StrZ As String) As String
    Dim lLen As Long
    lLen = StrLenW(StrPtr(StrZ))
    TrimZ = LeftB$(StrZ, lLen + lLen)
End Function

' note: this function behaves buggy when there is one NULL at the end of the string, and other NULLs within
' it is supposed to trim that NULL, but instead it trims to the first NULL that is found
' workaround: use a buffer that is larger by one character than you actually need
Public Function TrimZZ(StrZZ As String) As String
    Const ZZ As String = vbNullChar & vbNullChar
    Dim Idx As Long
    Idx = InStr(StrZZ, ZZ)
    If (Idx) Then
        TrimZZ = LeftB$(StrZZ, Idx + Idx - 2)
    Else
        Idx = InStr(StrZZ, vbNullChar)
        If (Idx) Then
            TrimZZ = LeftB$(StrZZ, Idx + Idx - 2) 'Rd
        Else
            TrimZZ = Trim$(StrZZ)
        End If
    End If
End Function

' note: this function behaves buggy when there is one NULL at the end of the string, and other NULLs within
' it is supposed to trim that NULL, but instead it trims to the first NULL that is found
' workaround: use a buffer that is larger by one character than you actually need
Public Function TrimZZ_2(StrZZ As String) As String
    Const ZZ As String = vbNullChar & vbNullChar
    Dim Idx As Long
    Do: Idx = InStrB(Idx + 1, StrZZ, ZZ)
    Loop While (Idx <> 0&) And ((Idx And 1&) = 0&)
    If Idx <> 0& Then
        TrimZZ_2 = LeftB$(StrZZ, Idx - 1&)
    Else
        Do: Idx = InStrB(Idx + 1, StrZZ, vbNullChar)
        Loop While (Idx <> 0&) And ((Idx And 1&) = 0&)
        If Idx <> 0& Then
            TrimZZ_2 = LeftB$(StrZZ, Idx - 1&)
        Else
            TrimZZ_2 = StrZZ
        End If
    End If
End Function

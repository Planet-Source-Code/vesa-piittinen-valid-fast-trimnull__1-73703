Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)

' trims all the NULLs at the start of the string
Public Function LTrimZ(Text As String) As String
    Static LH(0 To 5) As Long, IH(0 To 5) As Long, LPH As Long, IPH As Long
    Dim LA() As Long, LP As Long, IA() As Integer, IP As Long
    Dim I As Long, U As Long
    U = Len(Text) - 1
    If U >= 0 Then
        If LH(0) = 0 Then
            LH(0) = &H110001: LH(1) = 4: LH(4) = &H3FFFFFFF
            IH(0) = &H110001: IH(1) = 2
            LPH = VarPtr(LH(0))
            IPH = VarPtr(IH(0))
        End If
        LP = ArrPtr(LA)
        PutMem4 LP, LPH
        IH(3) = StrPtr(Text): IH(4) = U + 1
        IP = ArrPtr(IA)
        LH(3) = IP: LA(0) = IPH
        For I = 0 To U
            If IA(I) <> 0 Then Exit For
        Next I
        LA(0) = 0
        LH(3) = LP: LA(0) = 0
        If I = 0 Then
            LTrimZ = Text
        ElseIf I <= U Then
            LTrimZ = Right$(Text, U - I + 1)
        End If
    End If
End Function

' trims all NULLs from the end of the string
Public Function RTrimZ(Text As String) As String
    Static LH(0 To 5) As Long, IH(0 To 5) As Long, LPH As Long, IPH As Long
    Dim LA() As Long, LP As Long, IA() As Integer, IP As Long
    Dim I As Long, U As Long
    U = Len(Text) - 1
    If U >= 0 Then
        If LH(0) = 0 Then
            LH(0) = &H110001: LH(1) = 4: LH(4) = &H3FFFFFFF
            IH(0) = &H110001: IH(1) = 2
            LPH = VarPtr(LH(0))
            IPH = VarPtr(IH(0))
        End If
        LP = ArrPtr(LA)
        PutMem4 LP, LPH
        IH(3) = StrPtr(Text): IH(4) = U + 1
        IP = ArrPtr(IA)
        LH(3) = IP: LA(0) = IPH
        For I = U To 0 Step -1
            If IA(I) <> 0 Then Exit For
        Next I
        LA(0) = 0
        LH(3) = LP: LA(0) = 0
        If I = U Then
            RTrimZ = Text
        ElseIf I >= 0 Then
            RTrimZ = Left$(Text, I + 1)
        End If
    End If
End Function

' trim NULLs from the end of the string until non-null encountered OR until first dual NULL is encountered
' this is a bit fuzzy logic, but is somewhat a speed optimization for cases where NULL is wanted but dual-NULL is not
Public Function RTrimZZ(Text As String) As String
    Static LH(0 To 5) As Long, IH(0 To 5) As Long, LPH As Long, IPH As Long
    Dim LA() As Long, LP As Long, IA() As Integer, IP As Long
    Dim I As Long, N As Long, U As Long
    U = Len(Text) - 1
    If U >= 0 Then
        If LH(0) = 0 Then
            LH(0) = &H110001: LH(1) = 4: LH(4) = &H3FFFFFFF
            IH(0) = &H110001: IH(1) = 2
            LPH = VarPtr(LH(0))
            IPH = VarPtr(IH(0))
        End If
        LP = ArrPtr(LA)
        PutMem4 LP, LPH
        IH(3) = StrPtr(Text): IH(4) = U + 1
        IP = ArrPtr(IA)
        LH(3) = IP: LA(0) = IPH
        For I = 0 To U \ 2
            If IA(I) = 0 Then
                If I - 1 = N Then
                    I = N - 1
                    Exit For
                Else
                    N = I
                End If
            End If
            If IA(U - I) <> 0 Then I = U - I: Exit For
        Next I
        LA(0) = 0
        LH(3) = LP: LA(0) = 0
        If I = U Then
            RTrimZZ = Text
        ElseIf I >= 0 Then
            RTrimZZ = Left$(Text, I + 1)
        End If
    End If
End Function

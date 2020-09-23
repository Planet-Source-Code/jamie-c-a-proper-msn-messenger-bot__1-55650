Attribute VB_Name = "BitAnd"
Public Function BitAnd(ByVal String1 As String, ByVal String2 As String) As String
    Dim LenStr As Byte, LenStr2 As Byte
    LenStr = Len(String1)
    LenStr2 = Len(String2)
    If LenStr = 0 Or LenStr2 = 0 Then Exit Function


    If LenStr2 < LenStr Then
        String1 = Left(String1, LenStr2)
        LenStr = LenStr2
    Else
        String2 = Left(String2, LenStr)
    End If
    BitAnd = String(LenStr, vbNullChar)


    For LenStr2 = 1 To LenStr
        Mid(BitAnd, LenStr2, 1) = Chr(Asc(Mid(String1, LenStr2, 1)) Xor Asc(Mid(String2, LenStr2, 1)))
    Next
End Function


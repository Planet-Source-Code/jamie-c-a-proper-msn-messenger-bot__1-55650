Attribute VB_Name = "modStrings"
Public Function URLEncode(strData As String) As String

Dim intCount As Integer
Dim strBuffer As String
Dim strReturn As String

strReturn = strData

    For intCount = 1 To Len(strData)
    
        strBuffer = Mid(strData, intCount, 1)

        If Not strBuffer Like "[a-z,A-Z,0-9]" Then
        
            strReturn = Replace(strReturn, strBuffer, "%" & Hex(Asc(strBuffer)))
            
        End If
        
    Next intCount
    
    URLEncode = strReturn
    
End Function
Public Function URLDecode(strInput As String) As String

Dim strCodedChar  As String
Dim intBeginBy As Integer

intBeginBy = 1

Begin:

For bp1 = intBeginBy To Len(strInput)

If Mid(strInput, bp1, 1) = "%" Then

    strCodedChar = Mid(strInput, bp1 + 1, 1) & Mid(strInput, bp1 + 2, 1)
    
    strInput = Left(strInput, bp1 - 1) & Chr(Val("&H" & strCodedChar)) & Right(strInput, Len(strInput) - bp1 - 2)
    
    intBeginBy = bp1
    
    GoTo Begin
    
End If

Next bp1

URLDecode = strInput

End Function


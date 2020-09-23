VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Msn Bot Log....."
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock PRVSck 
      Index           =   0
      Left            =   240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Messenger 
      Left            =   240
      Top             =   3180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   2700
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtOutput 
      BackColor       =   &H0080C0FF&
      Height          =   3765
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Created By Jamie Chalker"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Visual Basic Msn Bot Example
'Coded By Jamie C
'Visit My Website: www.lillysoft.tk
Const strServer As String = "messenger.hotmail.com"
Const lngPort As Long = 1863
Dim SessionID As String
Dim AuthString As String
Dim strCurrentServer As String
Dim lngCurrentPort As Long
Dim intTrailid As Integer
Dim intConnState As Integer
Dim strLastSendCMD As String
Public NewSocks As Integer

Sub IncrementPrvTrailId(Socket)
   PrvTrailId(Socket) = PrvTrailId(Socket) + 1
End Sub
Public Function SData(Socket As Integer, Data As String)
    txtOutput = txtOutput & "Private Out: " & Data & vbNewLine
    PRVSck(Socket).SendData Data
    IncrementPrvTrailId (Socket)
End Function
Sub IncrementTrailID()

intTrailid = intTrailid + 1

End Sub

Sub IncrementState()

intConnState = intConnState + 1

End Sub

Sub ResetVars()

intConnState = 0
intTrailid = 1

End Sub

Public Sub ProcessData(strData As String)

strBuffer = strBuffer & strData

' MsgBox strBuffer

End Sub

Private Sub Form_Load()
Dim x As Integer
For x = 1 To StrSocket
    Load PRVSck(x)
Next
ResetVars
Messenger.Close
Messenger.Connect strServer, lngPort
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Winsock1.Close
    Messenger.Close
    End
End Sub
Public Function NextSock() As Integer
    NextSock = NewSocks
    NewSocks = NewSocks + 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub

Private Sub Messenger_Connect()

intConnState = 1
Messenger_DataArrival 0

End Sub
Public Sub dataout(Message As String)
    On Error Resume Next
    Messenger.SendData Message
    IncrementTrailID
    txtOutput = txtOutput & "DS Out: " & Message & vbNewLine
End Sub
Private Sub Messenger_DataArrival(ByVal bytesTotal As Long)

' This sub handles all incoming traffic from the
' Dispatch Server (DS) and Notification Server (NS)
'-----------------------------
Dim Linenum As Integer
Dim strRawData As String, strInput As String
Dim strHashParams As String
Dim strResponse As String
Dim SplitData As Variant
Dim varParams As Variant

Messenger.GetData strRawData, vbString

txtOutput = txtOutput & "DS In: " & strRawData & vbNewLine

If intConnState >= 6 Then

  For Linenum = 0 To UBound(Split(strRawData, vbCrLf))
  
    strInput = Split(strRawData, vbCrLf)(Linenum)
    SplitData = Split(strInput, " ")
    Select Case Left(strInput, 3)
   
    Case "CHL":
         'CHL - Challenge - MSN Sends This, You must reply with an MD5 Dump
        Dim strChallenge As String
        strChallenge = Replace(Split(strInput, " ")(2), vbCrLf, "")
        dataout "QRY " & intTrailid & " msmsgs@msnmsgr.com 32" & vbCrLf & MD5String(strChallenge & "Q1P7W2E4J9R8U3S5")
        
    Case "CHG":
        'CHG - Change - Your Status has changed, Time to Sync List
        dataout "SYN " & intTrailid & " 0" & vbCrLf
    
    Case "BLP":
        'BLP - N/A - Used to tell the server how to handle your messages
        If SplitData(1) = "BL" Then
            dataout "BLP " & intTrailid & " AL" & vbCrLf
        End If
    Case "RNG":
        temp = Split(SplitData(2), ":")
        sock = NextSock
        PRVSck(sock).Close
        PRVSck(sock).RemoteHost = temp(0)
        PRVSck(sock).RemotePort = temp(1)
        AuthString = SplitData(4)
        SessionID = SplitData(1)
        PRVSck(sock).Connect
        
    Case "LST":
        'Contact Lists
        If SplitData(3) >= 10 Then
        ElseIf SplitData(3) = 2 Then
          
        ElseIf SplitData(3) = "4" Or SplitData(3) = "3" Or SplitData(3) = "8" Then
            dataout "ADD " & intTrailid & " AL " & SplitData(1) & " " & SplitData(1) & vbCrLf
        End If
        
    Case "ADD":
        'Someone Added You
        If SplitData(2) = "RL" Then
             dataout "ADD " & intTrailid & " AL " & SplitData(3) & " " & SplitData(3) & vbCrLf
        End If
    End Select
    Next
End If

Select Case intConnState

    Case 1
        
            ' Handshake
            '-----------------------------
            
        strLastSendCMD = "VER " & intTrailid & " MSNP9 MSNP8 CVR0" & vbCrLf
    
        Messenger.SendData strLastSendCMD
        
        Call IncrementTrailID
        Call IncrementState
        
    Case 2
    
            ' Send client information to DS
            '-----------------------------

        If strRawData = strLastSendCMD Then
        
            strLastSendCMD = "CVR " & intTrailid & " 0x0413 winnt 5.2 i386 MSNMSGR 6.0.0268 MSMSGS " & StrUsername & vbCrLf
            
            Messenger.SendData strLastSendCMD
            
            Call IncrementTrailID
            Call IncrementState
            
        Else
        
            MsgBox "No support for this protocol."
            
        End If
        
        
        
    Case 3
    
    
            ' Send login name (xxx@xxx.xxx) to DS
            '-----------------------------
        
        strLastSendCMD = "USR " & intTrailid & " TWN I " & StrUsername & vbCrLf
        
        Messenger.SendData strLastSendCMD
        
        Call IncrementTrailID
        Call IncrementState
    
    
    
    Case 4
    
    
            ' Send password to DS or move to other server
            '-----------------------------

        If UCase$(Left$(strRawData, 4)) = "USR " Then
        

            ' Get the hash supplied by the DS:
            h = InStr(LCase$(strRawData), " lc")
            strHashParams = Right$(strRawData, Len(strRawData) - h)
            
            ' Start the SSL-procedure:
            strResponse = DoSSL(strHashParams)
            
            ' Pass authentication result back to the DS:
            strLastSendCMD = "USR " & CStr(intTrailid) & " TWN S " & strResponse & vbCrLf
            
            Messenger.SendData strLastSendCMD
            
            Call IncrementTrailID
            Call IncrementState
        
        ElseIf UCase$(Left(strRawData, 4)) = "XFR " Then
        
            ' Move to another server
            
            varParams = Split(strRawData, " ")
            strConnectionString = varParams(3)
            
            varParams = Split(strConnectionString, ":")
            strCurrentServer = varParams(0)
            lngCurrentPort = CLng(varParams(1))
            
            ResetVars
            
            Messenger.Close
            Messenger.Connect strCurrentServer, lngCurrentPort
        
        End If
        
        
        
    Case 5
    
    
            ' Authentication ok or failed?
            '-----------------------------
            
        If UCase$(Left$(strRawData, 4)) = "USR " Then
        dataout "CHG " & intTrailid & " NLN" & vbCrLf
        dataout "REA " & intTrailid & " " & StrUsername & " " & Replace(fname, " ", "%20") & vbCrLf
        Call IncrementState
        ElseIf UCase$(Left$(strRawData, 4)) = "911 " Then
        
MsgBox "Invalid password"
        
        End If
        
        
        
    Case 6
    
    
            ' Receive Hotmail Crap
            '-----------------------------
            
        If UCase$(Left$(strRawData, 4)) = "MSG " Then
        
            Messenger.SendData "CHG " & CStr(intTrailid) & " NLN" & vbCrLf
            
            Call IncrementTrailID
            Call IncrementState
            
        Else
        
            Call IncrementState
            GoTo LoginDone
            
        End If
        
        
        
    Case 7
    
        ' Continue With Bot Session
        '-----------------------------

LoginDone:

                    
            

End Select


'For debug purposes:
'-----------------------------

If intConnState <> 2 Then

    Debug.Print "S: > " & strRawData
    strRawData = ""

End If

If intConnState <> 4 And Len(strLastSendCMD) <> 0 Then

    Debug.Print "-  C: > " & strLastSendCMD
    
    If intConnState = 2 Or intConnState = 4 Then
    Else
        strLastSendCMD = ""
    End If
    
End If

End Sub

Private Sub PRVSck_Connect(Index As Integer)
    IncrementPrvTrailId (Index)
    SData Index, "ANS " & PrvTrailId(idex) & " " & StrUsername & " " & AuthString & " " & SessionID & vbCrLf
End Sub

Private Sub PRVSck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    PRVSck(Index).GetData strData
    txtOutput = txtOutput & strData & vbCrLf
    Select Case Left(strData, 3)
            Case "MSG":
            Bot.ParseMessage Index, strData
            Case "IRO":
            Bot.SocketNumber = Index
            Bot.SendMSG "Enter Response To New Convo ID"
    End Select
End Sub

Private Sub txtOutput_Change()
txtOutput.SelStart = Len(txtOutput.Text)
End Sub

Public Sub Winsock1_Close()

' Handle SSL connection
'-----------------------------------------------

    Layer = 0
    Winsock1.Close
    Set SecureSession = Nothing

End Sub

Public Sub Winsock1_Connect()

' Handle SSL connections
'-----------------------------------------------

    Set SecureSession = New clsCrypto
    Call SendClientHello(Winsock1)

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

' Decode SSL Information
' Passes result to the ProcessData() sub
'-----------------------------------------------

    'Parse each SSL Record
    Dim TheData As String
    Dim ReachLen As Long

    Do
    
        If SeekLen = 0 Then
            If bytesTotal >= 2 Then
                Winsock1.GetData TheData, vbString, 2
                SeekLen = BytesToLen(TheData)
                bytesTotal = bytesTotal - 2
            Else
                Exit Sub
            End If
        End If
        
        If bytesTotal >= SeekLen Then
            Winsock1.GetData TheData, vbString, SeekLen
            bytesTotal = bytesTotal - SeekLen
        Else
            Exit Sub
        End If
        
        
        Select Case Layer
            Case 0:
                ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
                CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
                Call IncrementRecv
                Call SendMasterKey(Winsock1)
            Case 1:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                    If VerifyMAC(TheData) Then Call SendClientFinish(Winsock1)
                Else
                    Winsock1.Close
                End If
             Case 2:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) = False Then Winsock1.Close
                Layer = 3
                
             Case 3:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) Then Call ProcessData(Mid(TheData, 17))
        End Select
    
        SeekLen = 0

    Loop Until bytesTotal = 0

End Sub

Function DoSSL(strChallenge As String) As String

' Handles the SSL part of the authentication
'-----------------------------------------------

    Dim varLines As Variant
    Dim varURLS As Variant
    
    Dim intCurCookie As Integer
    
    Dim strAuthInfo As String
    Dim StrHeader As String
    Dim strLoginServer As String
    Dim strLoginPage As String
    

    
    Dim colURLS As New Collection
    Dim colHeaders As New Collection


    
    'strChallenge = Replace(strChallenge, ",", "&")
    
'Connect to NEXUS:
'--------------------------------------------------
    strBuffer = ""
    
    Winsock1.Close
    Winsock1.Connect "nexus.passport.com", 443

    ' Wait for the SSL layer to be established:
    
    Do Until Layer = 3
        DoEvents
    Loop

    'Obtain login information from NEXUS:
    
    Call SSLSend(Winsock1, "GET /rdr/pprdr.asp HTTP/1.0" & vbCrLf & vbCrLf)
    
    Do Until InStr(1, strBuffer, vbCrLf & vbCrLf) <> 0
        DoEvents
    Loop
    
    Winsock1.Close
    
'--------------------------------------------------
'Done with NEXUS
    
    
    
'Begin processing data from NEXUS:
'--------------------------------------------------
    
    intCurCookie = 0
    varLines = Split(strBuffer, vbCrLf)
    
    ' Search for the header "PasswordURLs:"
    
        For intCount = LBound(varLines) To UBound(varLines)
        
            ' Add the values for "PasswordURLs:" to a collection:
            
            If Left$(CStr(varLines(intCount)), InStr(1, varLines(intCount), " ")) = "PassportURLs: " Then
                colHeaders.Add Right$(CStr(varLines(intCount)), Len(varLines(intCount)) - InStr(1, varLines(intCount), " ")), Left(varLines(intCount), InStr(1, varLines(intCount), " "))
                Exit For
            End If
            
        Next intCount
        
    
    varURLS = Split(colHeaders.Item("PassportURLs: "), ",")
    
    For intCount = LBound(varURLS) To UBound(varURLS)
        colURLS.Add Right(varURLS(intCount), Len(varURLS(intCount)) - InStr(1, varURLS(intCount), "=")), Left(varURLS(intCount), InStr(1, varURLS(intCount), "="))
    Next intCount
    
    'Get the server and page for logging in:

    strLoginServer = Left$(colURLS("DALogin="), InStr(1, colURLS("DALogin="), "/") - 1)
    strLoginPage = Right$(colURLS("DALogin="), Len(colURLS("DALogin=")) - InStr(1, colURLS("DALogin="), "/") + 1)
    
'--------------------------------------------------
'End processing
    

    
ConnectLogin:

'Connect to login server
'--------------------------------------------------

    strBuffer = ""
    
    ' Layer resembles the state of the SSL connection:
    Layer = 0
    
    Winsock1.Close
    Winsock1.Connect strLoginServer, 443

    ' Wait for the SSL layer to be established:
    
    Do Until Layer = 3
        DoEvents
    Loop

    StrHeader = "GET " & strLoginPage & " HTTP/1.1" & vbCrLf & _
                "Authorization: Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & Replace(StrUsername, "@", "%40") & ",pwd=" & URLEncode(StrPassword) & "," & strChallenge & _
                "User-Agent: MSMSGS" & vbCrLf & _
                "Host: loginnet.passport.com" & vbCrLf & _
                "Connection: Keep-Alive" & vbCrLf & _
                "Cache-Control: no-cache" & vbCrLf & vbCrLf

    Call SSLSend(Winsock1, StrHeader)

    ' Wait for the header to be recieved
    
    Do Until InStr(1, strBuffer, vbCrLf & vbCrLf) <> 0
        DoEvents
    Loop
    
    Dim strHeaderValue As String

    strHeaderValue = GetHeader("authentication-info:", strBuffer)
    
    If RequiresRedirect(strHeaderValue) = True Then
    
        strHeaderValue = GetHeader("location:", strBuffer)
        
        lngCharPos = InStr(strHeaderValue, "://")
        
        If (LCase$(Left$(strHeaderValue, lngCharPos - 1)) = "https") Then
        
            strLoginServer = Mid$(strHeaderValue, lngCharPos + 3, InStr(lngCharPos + 3, strHeaderValue, "/") - (lngCharPos + 3))
            strLoginPage = Right$(strHeaderValue, Len(strHeaderValue) - (InStr(lngCharPos + 3, strHeaderValue, "/") - 1))
            
            GoTo ConnectLogin
            
        End If
    
    Else
    
        DoSSL = ParseHash(strHeaderValue)
        Winsock1.Close
        Exit Function

    End If

'--------------------------------------------------
'Done with login server

End Function


Function GetHeader(StrHeader As String, strData As String) As String

' Returns the value of a header-property
'-----------------------------------------------

Dim intCount As Integer
Dim varLines As Variant
Dim lngCharPos As Long
Dim strCurHeader As String

varLines = Split(strData, vbCrLf)

For intCount = LBound(varLines) To UBound(varLines)

If Len(varLines(intCount)) = 0 Then Exit For

    strCurHeader = varLines(intCount)
    lngCharPos = InStr(strCurHeader, " ")
    
    If LCase(Left(strCurHeader, lngCharPos - 1)) = LCase(StrHeader) Then
        GetHeader = Right(strCurHeader, Len(strCurHeader) - lngCharPos)
        Exit Function
    End If
    

Next intCount

End Function

Function RequiresRedirect(strData As String) As Boolean

' Checks whether it's necessary to redirect to
' another server (using 'da-status' property)
'-----------------------------------------------

Dim intCount As Integer
Dim varProps As Variant
Dim lngCharPos As Long
Dim strCurItem As String
Dim strPropName As String
Dim strPropValue As String

lngCharPos = InStr(strData, " ")

If Left(strData, lngCharPos - 1) = "Passport1.4" Then

    strData = Right(strData, Len(strData) - lngCharPos)
    varProps = Split(strData, ",")
    
    For intCount = LBound(varProps) To UBound(varProps)
    
        strCurItem = varProps(intCount)
        lngCharPos = InStr(strCurItem, "=")
        
        strPropName = Left(strCurItem, lngCharPos - 1)
        strPropValue = Right(strCurItem, Len(strCurItem) - lngCharPos)
    
        If LCase$(strPropName) = "da-status" And LCase$(strPropValue) = "redir" Then
        
            RequiresRedirect = True
            Exit Function
            
        ElseIf LCase$(strPropName) = "da-status" And LCase$(strPropValue) = "success" Then
        
            RequiresRedirect = False
            Exit Function
        
        End If
        
    Next intCount

End If

End Function

Function ParseHash(StrHeader As String) As String

' Returns the hash (from-pp) if the login has
' completed succesfully.
'-----------------------------------------------

Dim intCount As Integer
Dim varProps As Variant
Dim lngCharPos As Long
Dim strCurItem As String
Dim strPropName As String
Dim strPropValue As String

    varProps = Split(StrHeader, ",")
    
    For intCount = LBound(varProps) To UBound(varProps)
    
        strCurItem = varProps(intCount)
        lngCharPos = InStr(strCurItem, "=")
        
        strPropName = Left(strCurItem, lngCharPos - 1)
        strPropValue = Right(strCurItem, Len(strCurItem) - lngCharPos)
    
        If LCase$(strPropName) = "from-pp" Then
        
            ParseHash = strPropValue
            'MsgBox ParseHash
            ParseHash = Left(ParseHash, Len(ParseHash) - 1)
            ParseHash = Right(ParseHash, Len(ParseHash) - 1)
            
            Exit Function
        
        End If
        
    Next intCount

End Function

'Coded And Created By Jamie C

Attribute VB_Name = "Bot"
'==============================
'Visual Basic Msn Bot Example
'Created By Jamie C.
'==============================
'Your Bots Email And Password Details Go Below
Public Const StrUsername As String = "Put Email Here"
Public Const StrPassword As String = "Put Password Here"
'Private Sockets To Load And Handle
Public Const StrSocket As Integer = 100
'Please Choose A Nickname For Your Bot (Plain Text)
Public Const StrName As String = "Fusion :: Type !menu to begin"

'Public Declarations
Public SocketNumber As Integer
Public PrvTrailId(StrSocket) As Integer
'Parse MSG Packet
Public Function ParseMessage(Socket As Integer, strData As String)
    On Error Resume Next
    Dim strUname As String
    Dim StrFName As String
    Dim Message As String
    Dim LineDat As Variant
    strUname = Split(strData, " ")(1)
    StrFName = Split(strData, " ")(2)
    Message = strData
    LineDat = Split(strData, vbCrLf)
    Message = LineDat(5)
    SocketNumber = Socket
    If Message <> "" And Message <> " " Then
        MessageFn strUname, Socket, Message
    End If
End Function

'Messages And Responses From Bot
'Majority Of Your Code Should Go Here
Private Function MessageFn(Email As String, Socket As Integer, Message As String)
    If Left(Message, 5) = "!menu" Then
        SendMSG "Menu: Your Menu Should Go Here. Code Created By Jamie C."
    End If
    If Left(messge, 5) = "!invite" Then
        InviteUser (Right(Message, Len(Message) - 6))
    End If
End Function
'SendMSG, Saves you alot of time instead of writing 3 lines of code
Public Function SendMSG(Message As String)
        Dim StrHeader As String
        StrHeader = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=" & "VERDANA" & "; EF=B; CO=" & "FF" & "; CS=0; PF=22" & vbCrLf & vbCrLf
        StrHeader = StrHeader & Message
        StrHeader = "MSG " & PrvTrailId(SocketNumber) & " N " & Len(StrHeader) & vbCrLf & StrHeader
        frmMain.SData SocketNumber, StrHeader
End Function
'Invite User Function.
Public Function InviteUser(Email As String)
frmMain.SData SocketNumber, "CAL " & PrvTrailId(SocketNumber) & " " & Email & vbCrLf
End Function

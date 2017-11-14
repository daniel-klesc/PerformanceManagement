VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public destination As String
Public Level As Integer
Public category As Variant

Public subjMsgLenght As Integer
Public mailAddress As String

Public Function log(lvl As Integer, dest As String, Message As MSG)
    If Not lvl < Me.Level Then
        If dest = Me.destination Then
            save Message, lvl
        End If
    End If
End Function

Public Function save(Message As MSG, lvl As Integer)
    Dim oApp As Object
    Dim oMail As Object
    
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    With oMail
        .To = mailAddress
        .Subject = Message.source_ & " " & toMsgType(lvl) & " " & Left(Message.text_, subjMsgLenght)
        .Body = Now & vbCrLf & Message.source_ & vbCrLf & toMsgType(lvl) & vbCrLf & Message.text_
        .Send
    End With
    
End Function

Public Function init(str_name As String, lvl As Integer, str_dest As String)
    Name = str_name
    Level = lvl
    destination = str_dest
    subjMsgLenght = 10
End Function

Private Function toMsgType(lvl As Integer) As String
    Select Case lvl
        Case 1:
            toMsgType = "TRACE"
        Case 2:
            toMsgType = "DEBUG"
        Case 3:
            toMsgType = "INFO"
        Case 4:
            toMsgType = "WARN"
        Case 5:
            toMsgType = "ERROR"
    End Select
End Function


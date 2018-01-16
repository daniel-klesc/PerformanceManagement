VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String
Public destination As String
Public Level As Integer
Public category As Variant


Public wsName As String
Public logFilePath As String
Public firstCell As String

Public is_same_app As Boolean

Private logWB As Workbook
Private logIsOpen As Boolean

Private DATE_TIME_OFFSET As Integer
Private TYPE_OFFSET As Integer
Private SOURCE_OFFSET As Integer
Private MESSAGE_OFFSET As Integer
Private RESULT_TYPE As Integer


Public Function log(lvl As Integer, dest As String, message As MSG)
    If Not lvl < Me.Level Then
        If dest = Me.destination Then
            save message, lvl
        End If
    End If
End Function



Private Function save(message As MSG, lvl As Integer)
    Dim ws As Worksheet
    Dim newRowIndex As Long
    Dim actualRng As Range
    
    If Not logIsOpen Then
        Set logWB = openFile
    End If
    
    Set ws = logWB.Worksheets(wsName)
        
    newRowIndex = ws.UsedRange.Rows.Count + 1
    
    Set actualRng = ws.Cells(newRowIndex, ws.Range(firstCell).Column)
    
    actualRng.Offset(0, DATE_TIME_OFFSET).Value = CStr(Now)
    actualRng.Offset(0, TYPE_OFFSET).Value = toMsgType(lvl)
    actualRng.Offset(0, SOURCE_OFFSET).Value = message.source_
    actualRng.Offset(0, MESSAGE_OFFSET).Value = message.text_
    
    If message.params.Count <> 0 Then
        actualRng.Offset(0, RESULT_TYPE).Value = message.params(1)
    End If
    
End Function

Public Function init(str_name As String, lvl As Integer, str_dest As String)
    name = str_name
    Level = lvl
    destination = str_dest
    
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

Private Function openFile() As Workbook
    Dim lApp As Application
    
    Set lApp = CreateObject("Excel.Application")
    lApp.Visible = False
    
    On Error GoTo ERR_CANT_OPEN
        lApp.Workbooks.Open logFilePath
        Set openFile = lApp.Workbooks(1)
    On Error GoTo 0
    
    logIsOpen = True
    
    Exit Function
ERR_CANT_OPEN:
    lApp.Quit
    Err.raise Err.Number, Err.source, "Logovací soubor se nepodaøilo otevøít. Zkontrolujte zda je následující cesta správná: " & logFilePath
End Function

Public Function closeFile()
    Dim lApp As Application
    
    If logIsOpen Then
        Set lApp = logWB.Parent
        logWB.Close SaveChanges:=True
        lApp.Quit
        logIsOpen = False
    End If
End Function

Private Sub Class_Initialize()
    wsName = "db.log"
    firstCell = "A2"
    logIsOpen = False
    
    DATE_TIME_OFFSET = 0
    TYPE_OFFSET = 1
    SOURCE_OFFSET = 2
    MESSAGE_OFFSET = 3
    RESULT_TYPE = 4
End Sub


Private Sub Class_Terminate()
    If logIsOpen Then
        closeFile
    End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogFileReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private WithEvents pSpecs As SpecSuite
Attribute pSpecs.VB_VarHelpID = -1
Private message As MSG

' ============================================= '
' Public Methods
' ============================================= '

''
' Listen to given SpecSuite
'
' @method ListenTo
' @param {SpecSuite} Specs
''
Public Sub ListenTo(specs As SpecSuite)
    
    'Debug.Print "===" & IIf(specs.Description <> "", " " & specs.Description & " ===", "")
    Set pSpecs = specs
End Sub

Public Function Done(suiteName As String, noOfPassed As Integer, noOfTotal As Integer)
    
    Set message = New MSG
    log4VBA.info log4VBA.TEST_DESTINATION, message.source(suiteName).text(CStr(noOfPassed) & " of " & CStr(noOfTotal) & " tests has passed.").addParams(IIf(noOfPassed = noOfTotal, "PASS", "FAIL"))
End Function


Private Sub pSpecs_Result(spec As SpecDefinition)
    Set message = New MSG
    If spec.result = Fail Then
        log4VBA.info log4VBA.TEST_DESTINATION, message.source(pSpecs.Description).text(IIf(spec.Id <> "", spec.Id & " ", "No of elements in collection ") & spec.Expectations(1).FailureMessage).addParams("FAIL")
    End If
End Sub


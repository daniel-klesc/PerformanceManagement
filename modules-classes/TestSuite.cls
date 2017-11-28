VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private reporter As LogFileReporter
Private suitName_ As String


Private testCases As Collection

Private Property Get noOfPassed() As Integer
    Dim testCase As TestCaseMDLGeneral
    Dim counter As Integer
    
    counter = 0
    
    For Each testCase In testCases
        If testCase.result Then
            counter = counter + 1
        End If
    Next testCase
    
    noOfPassed = counter
End Property

Public Property Let suitName(name As String)
    Me.suitName_ = name
End Property

Public Property Get suitName() As String
    suitName = suitName_
End Property


Public Function addTestCase(setting As SettingsMDLTest)
    Dim testCase As TestCaseMDLGeneral
    
    Set setting.reporter = Me.reporter
    
    Set testCase = New TestCaseMDLGeneral
    testCase.init setting
    
    testCases.add testCase, testCase.testName
    
End Function


Public Function runTests()
    Dim testCase As TestCaseMDLGeneral

    For Each testCase In testCases
        testCase.runTest
    Next testCase
    
    reporter.Done suitName, noOfPassed, testCases.Count
    
    
End Function

Private Sub Class_Initialize()
    Set testCases = New Collection
    Set reporter = New LogFileReporter
    
    
    suitName_ = "MDLTestSuiteGeneric"
End Sub



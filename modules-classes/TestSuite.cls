VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "testSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private reporter As LogFileReporter
Private suitName As String


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

Public Function init()

    Dim logger As New LoggerFile
    Dim testSettings As SettingsMDLTest
    Dim testCase As TestCaseMDLGeneral
    Dim MDLObject As MDLBINProdLine
    
    log4VBA.init
    logger.init suitName, log4VBA.INF, log4VBA.TEST_DESTINATION
    logger.logFilePath = "C:\Users\czJiRost\Desktop\log.xlsx"
    log4VBA.add_logger logger
    
    Set testCase = New TestCaseMDLGeneral
    Set testSettings = New SettingsMDLTest
    
    Set MDLObject = New MDLBINProdLine
    MDLObject.obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    MDLObject.obj_data_provider.str_path = "C:\Users\czJiRost\Desktop\"
    MDLObject.obj_data_provider.str_file_name = "test.xlsx"
    MDLObject.obj_data_provider.STR_WS_NAME = "data"
    Set testSettings.checkedMDL = MDLObject
    
    Set MDLObject = New MDLBINProdLine
    MDLObject.obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    MDLObject.obj_data_provider.str_path = "C:\Users\czJiRost\Desktop\"
    MDLObject.obj_data_provider.str_file_name = "test2.xlsx"
    MDLObject.obj_data_provider.STR_WS_NAME = "data"
    Set testSettings.testedMDL = MDLObject
    
    Set testSettings.reporter = reporter
    testSettings.testName = "Check MDLBINProdLine"
    
    Set testSettings.testedMDL = MDLObject
    
    testCase.init testSettings
    
    testCases.add testCase
    
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
    
    
    suitName = "MDLTestSuite"
End Sub

Private Sub Class_Terminate()
    log4VBA.remove_logger suitName
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCaseMDLGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String
Public testName As String

Private Enum loadStatus
    Checked
    Tested
End Enum

Private testResult As Boolean

Private switch As loadStatus


Private specs As SpecSuite
Private reporter As LogFileReporter


Private testCollection As Collection
Private checkedCollection As Collection



Private testedMDL As Object
Private checkedMDL As Object


Private Function loadCollections()
    
    If checkedMDL Is Nothing Then
        Err.raise 1, , "Checked MDL is not set, be sure to run init before runTest."
    ElseIf testedMDL Is Nothing Then
        Err.raise 2, , "Tested MDL is not set, be sure to run init before runTest."
    ElseIf reporter Is Nothing Then
        Err.raise 3, , "LogFileReporter is not set, be sure to run init before runTest."
    End If

    checkedMDL.add_listener Me
    testedMDL.add_listener Me
    
    switch = Checked
    checkedMDL.load
    
    switch = Tested
    testedMDL.load
    
End Function

Private Function testCollections()
    Dim counter As Long
    
    specs.Description = testName
    
    specs.It("Models should have same number of elements").Expect(checkedCollection.Count).ToEqual testCollection.Count
    
    If specs.FailedSpecs.Count <> 0 Then
        Exit Function
    End If
    
    For counter = 1 To checkedCollection.Count
        specs.It("All elements in both collection have to be same.", "Row no. " & CStr(counter + 1)).Expect(checkedCollection(counter).toString).ToEqual testCollection(counter).toString
        If specs.FailedSpecs.Count <> 0 Then
            Exit Function
        End If
        
    Next counter
    
    If specs.FailedSpecs.Count = 0 Then
        testResult = True
    End If
    
    
End Function


'main test function

Public Function runTest()
    
    loadCollections
    testCollections
End Function


Public Function init(settingObject As SettingsMDLTest)
    Set reporter = settingObject.reporter
    Set checkedMDL = settingObject.checkedMDL
    Set testedMDL = settingObject.testedMDL
    testName = settingObject.testName
    reporter.ListenTo specs
End Function

Public Property Get result() As Boolean
    result = testResult
End Property


'
' Implementation of mandatory methods for ModelListener
'

Public Function loading_data_is_about_to_start(providerInfo As FileExcelDataProviderInfo)

End Function

Public Function loading_data_has_finished(providerInfo As FileExcelDataProviderInfo)

End Function

Public Function process_record(dbRecord As Object)
    If switch = Checked Then
        checkedCollection.add dbRecord
    Else
        testCollection.add dbRecord
    End If
End Function


'
' Constructor :)
'

Private Sub Class_Initialize()
    str_module = "TestCaseMDLGeneral"
    
    testResult = False
    
    Set testCollection = New Collection
    Set checkedCollection = New Collection
    Set specs = New SpecSuite
End Sub



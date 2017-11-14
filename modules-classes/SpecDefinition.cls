VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' SpecDefinition v2.0.0-beta
' (c) Tim Hall - https://github.com/VBA-tools/VBA-TDD
'
' Collection of expectations for verifying spec
'
' @class SpecDefinition
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Events and Properties
' --------------------------------------------- '

Public Description As String
Public Id As String
Public Expectations As VBA.Collection

Public Suite As SpecSuite

Public Property Get result() As SpecResultType
    If Me.Expectations.Count = 0 Then
        result = SpecResultType.Pending
    Else
        result = SpecResultType.Pass
        
        Dim Expectation As SpecExpectation
        For Each Expectation In Me.Expectations
            If Not Expectation.Passed Then
                result = SpecResultType.Fail
                Exit For
            End If
        Next Expectation
    End If
End Property

Public Property Get FailedExpectations() As Collection
    Dim Filtered As New Collection
    Dim Expectation As SpecExpectation
    For Each Expectation In Me.Expectations
        If Not Expectation.Passed Then
            Filtered.add Expectation
        End If
    Next Expectation
    
    Set FailedExpectations = Filtered
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Create a new expectation to test the given actual value
'
' @method Expect
' @param {Variant} Actual value to test
' @return {SpecExpectation}
''
Public Function Expect(Optional Actual As Variant) As SpecExpectation
    Dim Expectation As New SpecExpectation
    
    If VBA.VarType(Actual) = VBA.vbObject Then
        Set Expectation.Actual = Actual
    Else
        Expectation.Actual = Actual
    End If
    Me.Expectations.add Expectation
    
    Set Expect = Expectation
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Sub Class_Initialize()
    Set Me.Expectations = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Me.Suite.SpecDone Me
End Sub


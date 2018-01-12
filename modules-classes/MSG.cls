VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public source_ As String
Public text_ As String
Public params As Collection


Public Function source(str_module As String, Optional str_function As String = "") As MSG
    If str_function <> "" Then
        Me.source_ = str_module & log4VBA.DEFAULT_SEPARATOR & str_function
    Else
        Me.source_ = str_module
    End If
    Set source = Me
End Function

Public Function text(str_text As String) As MSG
    Me.text_ = str_text
    Set text = Me
End Function


Public Function addParams(var_params As Variant, Optional paramName As String = "") As MSG
    Me.params.add var_params, paramName
    Set addParams = Me
End Function

Private Sub Class_Initialize()
    Set params = New Collection
End Sub

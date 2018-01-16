VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIAdditionalDataClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_date As String
Public str_time As String

Public str_shift As String

Public str_process As String
Public str_process_step As String

Public str_area As String

Public str_building As String

Public str_process_status As String

Public str_time_status As String

Public str_task_status As String

Public int_results_num As Integer

Public str_shift_type As String
Public str_datetime_shift_separator As String

Public is_ready_for_update As Boolean


Public Property Get str_datetime_shift() As String
    Dim str_date_shift_start As String ' date when shift actually started
    Dim int_hour As Integer

    int_hour = Hour(str_time)

    If int_hour >= 0 And int_hour < 7 Then
        str_date_shift_start = DateAdd("d", -1, str_date)
    Else
        str_date_shift_start = str_date
    End If

    str_datetime_shift = str_date_shift_start & str_datetime_shift_separator & str_shift_type
End Property

Public Function add_result_pallet(obj_add_data As KPIAdditionalData)
        int_results_num = int_results_num + 1
End Function

Public Function reset()
    int_results_num = 0
End Function


Private Sub Class_Initialize()
    int_results_num = 0
End Sub

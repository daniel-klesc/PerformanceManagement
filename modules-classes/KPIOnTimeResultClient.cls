VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimeResultClient"
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

Public str_kpi As String
'Public int_kpi_target As Integer
'Public int_kpi_actual As Integer
Public dbl_kpi_target As Double
Public dbl_kpi_actual As Double

'Public int_task_target As Integer
'Public int_task_actual As Integer
Public dbl_task_target As Double
Public dbl_task_actual As Double

Public int_tasks_num As Integer

Public str_building As String

Public str_datetime_shift As String


Private int_good_results_num As Integer
Private int_bad_results_num As Integer

Public Function add_result_pallet(obj_result_pallet As KPIOnTimeResultPallet)
    If obj_result_pallet.str_kpi_status = "True" Then
        int_good_results_num = int_good_results_num + 1
    Else
        int_bad_results_num = int_bad_results_num + 1
    End If
    
    int_tasks_num = int_tasks_num + 1
End Function

Public Function calculate()
    If int_good_results_num > 0 Or int_bad_results_num > 0 Then
        dbl_kpi_actual = int_good_results_num / (int_good_results_num + int_bad_results_num)
    End If
End Function

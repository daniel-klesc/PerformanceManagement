VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLAdditionalDataClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLAdditionalDataClient"

Public STR_DB_TYPE_FILE_EXCEL As String

Public INT_OFFSET_DATE As Integer
Public INT_OFFSET_TIME As Integer
Public INT_OFFSET_SHIFT As Integer

Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_PROCESS_STEP As Integer

Public INT_OFFSET_AREA As Integer

Public INT_OFFSET_BUILDING As Integer

Public INT_OFFSET_PROCESS_STATUS As Integer
Public INT_OFFSET_TIME_STATUS As Integer

Public INT_OFFSET_TASK_STATUS As Integer
Public INT_OFFSET_TASK_ACTUAL_VALUE As Integer

Public INT_OFFSET_DATETIME_SHIFT As Integer

Public obj_multi_data_provider As MultiFileExcelDataProvider
Public obj_current_data_provider_info As FileExcelDataProviderInfo
Public str_static_data_provider_id As String
    ' listeners must implement:
      ' process_record obj_db_data_process: DBDataProcess
Public col_listeners As Collection
    ' listeners must implement:
Public str_id As String
      ' loading_data_is_about_to_start obj_provider_info: FileExcelDataProviderInfo
'Public col_provider_info_listeners As Collection

Public str_datetime_shift_separator As String

Private Sub Class_Initialize()
    STR_DB_TYPE_FILE_EXCEL = "file:excel"

    INT_OFFSET_DATE = 0
    INT_OFFSET_TIME = 1
    INT_OFFSET_SHIFT = 2
    INT_OFFSET_PROCESS_ID = 3
    INT_OFFSET_PROCESS_STEP = 4
    INT_OFFSET_AREA = 5
    INT_OFFSET_BUILDING = 6
    INT_OFFSET_PROCESS_STATUS = 7
    INT_OFFSET_TIME_STATUS = 8
    INT_OFFSET_TASK_STATUS = 9
    INT_OFFSET_TASK_ACTUAL_VALUE = 10
    INT_OFFSET_DATETIME_SHIFT = 11
    
    str_static_data_provider_id = "static"
    Set col_listeners = New Collection
    
    str_datetime_shift_separator = "-"
    
    str_id = str_module

End Sub

Public Function set_clear_data()
    obj_multi_data_provider.set_clear_data_before_close_status True
End Function

Public Function reset_clear_data()
    obj_multi_data_provider.set_clear_data_before_close_status False
End Function

Public Function load_static()
    obj_multi_data_provider.add_listener Me
    obj_multi_data_provider.load_data_static
    obj_multi_data_provider.remove_listener str_id
End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_add_data_client As KPIAdditionalDataClient
    Dim obj_listener As Object

    ' fill data in a row
    Set obj_add_data_client = New KPIAdditionalDataClient
    
      ' date and time
    'obj_kpi_on_time_result_client.str_date = Format(rg_record.Offset(0, INT_OFFSET_DATE).Value, "DD.MM.YYYY")
    obj_add_data_client.str_date = Format(rg_record.Offset(0, INT_OFFSET_DATE).Value, "D.M.YYYY")
    obj_add_data_client.str_time = Format(rg_record.Offset(0, INT_OFFSET_TIME).Value, "h:mm:ss")
      ' shift
    obj_add_data_client.str_shift = rg_record.Offset(0, INT_OFFSET_SHIFT).Value
      ' process
    obj_add_data_client.str_process = rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value
    obj_add_data_client.str_process_step = rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value
    
    obj_add_data_client.str_building = rg_record.Offset(0, INT_OFFSET_BUILDING).Value
    obj_add_data_client.str_task_status = rg_record.Offset(0, INT_OFFSET_TASK_STATUS).Value
    obj_add_data_client.int_results_num = rg_record.Offset(0, INT_OFFSET_TASK_ACTUAL_VALUE).Value
    obj_add_data_client.str_datetime_shift_separator = str_datetime_shift_separator
    obj_add_data_client.str_shift_type = hndl_master_shift.find_type(Hour(obj_add_data_client.str_time))
    
    For Each obj_listener In col_listeners
        obj_listener.add_data_client_process_record obj_add_data_client
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        'obj_listener.loading_data_is_about_to_start obj_provider_info
        obj_listener.add_data_client_loading_data_is_about_to_start obj_provider_info
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        obj_listener.add_data_client_loading_data_has_finished obj_provider_info
    Next
End Function

Public Function save_record_static(obj_additional_data_client As KPIAdditionalDataClient)
    Dim obj_data_provider As FileExcelDataProvider
    Dim rg_next_data As Range

    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_static_data_provider_id)
    On Error GoTo INFO_PROVIDER_IS_NOT_OPEN
    Set rg_next_data = obj_data_provider.next_data
    On Error GoTo 0

    save_record obj_additional_data_client, rg_next_data
    
    Exit Function
INFO_PROVIDER_IS_NOT_OPEN:
    obj_data_provider.int_used_range_offset = 2
    obj_data_provider.bool_clear_after_open_status = True
    obj_data_provider.open_data
    obj_data_provider.bool_clear_after_open_status = False
    Set rg_next_data = obj_data_provider.next_data
    Resume Next
End Function

Public Function save_record(obj_additional_data_client As KPIAdditionalDataClient, rg_record As Range)
    ' date and time
    rg_record.Offset(0, INT_OFFSET_DATE).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_DATE).Value = Format(obj_additional_data_client.str_date, "DD.MM.YY")
    rg_record.Offset(0, INT_OFFSET_TIME).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_TIME).Value = Format(obj_additional_data_client.str_time, "hh:mm")
    ' shift
    rg_record.Offset(0, INT_OFFSET_SHIFT).Value = obj_additional_data_client.str_shift
    ' process
    rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value = obj_additional_data_client.str_process
    rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value = obj_additional_data_client.str_process_step
    
    rg_record.Offset(0, INT_OFFSET_BUILDING).Value = obj_additional_data_client.str_building
    
    rg_record.Offset(0, INT_OFFSET_TASK_STATUS).Value = obj_additional_data_client.str_task_status
    rg_record.Offset(0, INT_OFFSET_TASK_ACTUAL_VALUE).Value = obj_additional_data_client.int_results_num
    
    rg_record.Offset(0, INT_OFFSET_DATETIME_SHIFT).Value = obj_additional_data_client.str_datetime_shift
    
End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener, obj_listener.str_id
End Function

Public Function remove_listener(str_id As String)
    col_listeners.Remove str_id
End Function


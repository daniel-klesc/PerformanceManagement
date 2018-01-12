VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIOnTimeResultClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLKPIOnTimeResultClient"

Public STR_DB_TYPE_FILE_EXCEL As String

Public INT_OFFSET_DATE As Integer
Public INT_OFFSET_TIME As Integer
Public INT_OFFSET_SHIFT As Integer

Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_PROCESS_STEP As Integer

Public INT_OFFSET_KPI_ID As Integer

Public INT_OFFSET_KPI_TARGET As Integer
Public INT_OFFSET_KPI_ACTUAL As Integer

Public INT_OFFSET_TASK_TARGET As Integer
Public INT_OFFSET_TASK_ACTUAL As Integer

Public INT_OFFSET_TASK_NUM As Integer
Public INT_OFFSET_BUILDING As Integer

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
    
    INT_OFFSET_KPI_ID = 5
    
    INT_OFFSET_KPI_TARGET = 6
    INT_OFFSET_KPI_ACTUAL = 7
    
    INT_OFFSET_TASK_TARGET = 8
    INT_OFFSET_TASK_ACTUAL = 9
    
    INT_OFFSET_TASK_NUM = 10
    INT_OFFSET_BUILDING = 11
    
    INT_OFFSET_DATETIME_SHIFT = 12
        
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

'Public Function load_data_for_interval(obj_date_start As Date)
'    Dim obj_data_provider As FileExcelDataProvider
'    Dim obj_data_provider_util As FileExcelDataProviderUtil
'    Dim str_provider_id As String
'
'    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
'    obj_data_provider_util.str_datetime = obj_date_start
'    str_provider_id = obj_data_provider_util.str_provider_id
'    On Error GoTo INFO_PROVIDER_NOT_FOUND
'    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_provider_id)
'    On Error GoTo 0
'
'    obj_data_provider.bool_clear_before_close_status = True
'    On Error GoTo INFO_FILE_NOT_EXISTS
'    obj_data_provider.load_data
'    On Error GoTo 0
'    obj_multi_data_provider.remove_provider obj_data_provider
'    Exit Function
'INFO_PROVIDER_NOT_FOUND:
'    Set obj_data_provider = _
'        obj_multi_data_provider.create_data_provider_read_mode( _
'            str_provider_id, obj_data_provider_util.retrieve_creation_file_name)
'    obj_data_provider.bool_save_mode_on = True ' to be able save the file as empty one when loading is finished
'    Resume Next
'INFO_FILE_NOT_EXISTS:
'    ' file not exists yet => remove provider
'    obj_multi_data_provider.remove_provider obj_data_provider
'End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_kpi_on_time_result_client As KPIOnTimeResultClient
    Dim obj_listener As Object

    ' fill data in a row
    Set obj_kpi_on_time_result_client = New KPIOnTimeResultClient
    
      ' date and time
    'obj_kpi_on_time_result_client.str_date = Format(rg_record.Offset(0, INT_OFFSET_DATE).Value, "DD.MM.YYYY")
    obj_kpi_on_time_result_client.str_date = Format(rg_record.Offset(0, INT_OFFSET_DATE).Value, "D.M.YYYY")
    obj_kpi_on_time_result_client.str_time = Format(rg_record.Offset(0, INT_OFFSET_TIME).Value, "h:mm:ss")
      ' shift
    obj_kpi_on_time_result_client.str_shift = rg_record.Offset(0, INT_OFFSET_SHIFT).Value
      ' process
    obj_kpi_on_time_result_client.str_process = rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value
    obj_kpi_on_time_result_client.str_process_step = rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value
      ' kpi
    obj_kpi_on_time_result_client.str_kpi = rg_record.Offset(0, INT_OFFSET_KPI_ID).Value
    obj_kpi_on_time_result_client.dbl_kpi_target = rg_record.Offset(0, INT_OFFSET_KPI_TARGET).Value
    obj_kpi_on_time_result_client.dbl_kpi_actual = rg_record.Offset(0, INT_OFFSET_KPI_ACTUAL).Value
      ' task
    obj_kpi_on_time_result_client.dbl_task_target = rg_record.Offset(0, INT_OFFSET_TASK_TARGET).Value
    obj_kpi_on_time_result_client.dbl_task_actual = rg_record.Offset(0, INT_OFFSET_TASK_ACTUAL).Value
    obj_kpi_on_time_result_client.int_tasks_num = rg_record.Offset(0, INT_OFFSET_TASK_NUM).Value
      ' place
    obj_kpi_on_time_result_client.str_building = rg_record.Offset(0, INT_OFFSET_BUILDING).Value
    'obj_kpi_on_time_result_client.str_datetime_shift = rg_record.Offset(0, INT_OFFSET_DATETIME_SHIFT).Value
    obj_kpi_on_time_result_client.str_datetime_shift_separator = str_datetime_shift_separator
    obj_kpi_on_time_result_client.str_shift_type = hndl_master_shift.find_type(Hour(obj_kpi_on_time_result_client.str_time))
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_result_client_process_record obj_kpi_on_time_result_client
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        'obj_listener.loading_data_is_about_to_start obj_provider_info
        obj_listener.kpi_result_client_loading_data_is_about_to_start obj_provider_info
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_result_client_loading_data_has_finished obj_provider_info
    Next
End Function

'Public Function save_record_dynamic(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
'    Dim obj_data_provider As FileExcelDataProvider
'    Dim obj_data_provider_util As FileExcelDataProviderUtil
'    'Dim str_provider_id As String
'
'    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
'    obj_data_provider_util.str_datetime = obj_kpi_on_time_result_pallet.str_date & " " & obj_kpi_on_time_result_pallet.str_time
'    'str_provider_id = obj_data_provider_util.str_provider_id
'    On Error GoTo INFO_PROVIDER_NOT_FOUND
'    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(obj_data_provider_util.str_provider_id)
'    On Error GoTo 0
'
'    save_record obj_kpi_on_time_result_pallet, obj_data_provider.next_data
'    Exit Function
'INFO_PROVIDER_NOT_FOUND:
'    Set obj_data_provider = _
'        obj_multi_data_provider.create_data_provider_save_mode( _
'            obj_data_provider_util.str_provider_id, _
'            obj_data_provider_util.retrieve_creation_file_name, _
'            obj_data_provider_util.retrieve_creation_tmpl_file_name)
'    ' open file to start saving data
'    obj_data_provider.open_data
'    Resume Next
'End Function

Public Function save_record_static(obj_kpi_on_time_result_client As KPIOnTimeResultClient)
    Dim obj_data_provider As FileExcelDataProvider
    Dim rg_next_data As Range

    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_static_data_provider_id)
    On Error GoTo INFO_PROVIDER_IS_NOT_OPEN
    Set rg_next_data = obj_data_provider.next_data
    On Error GoTo 0

    save_record obj_kpi_on_time_result_client, rg_next_data
    
    Exit Function
INFO_PROVIDER_IS_NOT_OPEN:
    obj_data_provider.int_used_range_offset = 2
    obj_data_provider.bool_clear_after_open_status = True
    obj_data_provider.open_data
    obj_data_provider.bool_clear_after_open_status = False
    Set rg_next_data = obj_data_provider.next_data
    Resume Next
End Function

Public Function save_record(obj_kpi_on_time_result_client As KPIOnTimeResultClient, rg_record As Range)
    ' date and time
    rg_record.Offset(0, INT_OFFSET_DATE).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_DATE).Value = Format(obj_kpi_on_time_result_client.str_date, "DD.MM.YY")
    rg_record.Offset(0, INT_OFFSET_TIME).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_TIME).Value = Format(obj_kpi_on_time_result_client.str_time, "hh:mm")
    ' shift
    rg_record.Offset(0, INT_OFFSET_SHIFT).Value = obj_kpi_on_time_result_client.str_shift
    ' process
    rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value = obj_kpi_on_time_result_client.str_process
    rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value = obj_kpi_on_time_result_client.str_process_step
    ' kpi
    rg_record.Offset(0, INT_OFFSET_KPI_ID).Value = obj_kpi_on_time_result_client.str_kpi
    rg_record.Offset(0, INT_OFFSET_KPI_TARGET).Value = obj_kpi_on_time_result_client.dbl_kpi_target
    rg_record.Offset(0, INT_OFFSET_KPI_ACTUAL).Value = obj_kpi_on_time_result_client.dbl_kpi_actual
    ' task
    rg_record.Offset(0, INT_OFFSET_TASK_TARGET).Value = obj_kpi_on_time_result_client.dbl_task_target
    rg_record.Offset(0, INT_OFFSET_TASK_ACTUAL).Value = obj_kpi_on_time_result_client.dbl_task_actual
    rg_record.Offset(0, INT_OFFSET_TASK_NUM).Value = obj_kpi_on_time_result_client.int_tasks_num
    ' place
    rg_record.Offset(0, INT_OFFSET_BUILDING).Value = obj_kpi_on_time_result_client.str_building
    ' datetime id
    rg_record.Offset(0, INT_OFFSET_DATETIME_SHIFT).Value = obj_kpi_on_time_result_client.str_datetime_shift
End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener, obj_listener.str_id
End Function

Public Function remove_listener(str_id As String)
    col_listeners.Remove str_id
End Function

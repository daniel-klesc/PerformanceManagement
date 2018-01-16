VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLAdditionalData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLAdditionalData"

Public STR_DB_TYPE_FILE_EXCEL As String

Public INT_OFFSET_PALLET As Integer
Public INT_OFFSET_DATE As Integer
Public INT_OFFSET_TIME As Integer
Public INT_OFFSET_BUILDING As Integer
Public INT_OFFSET_PROCESS_NAME As Integer
Public INT_OFFSET_PROCESS_STEP As Integer
Public INT_OFFSET_TASK_STATUS As Integer

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

' tracking management
Public STR_PROCESSED_FILE_PATH_NEXT_MODULE As String
Public STR_PROCESSED_FILE_NAME_NEXT_MODULE As String


Private Sub Class_Initialize()
    STR_DB_TYPE_FILE_EXCEL = "file:excel"

    INT_OFFSET_PALLET = 0
    INT_OFFSET_DATE = 1
    INT_OFFSET_TIME = 2
    INT_OFFSET_BUILDING = 3
    INT_OFFSET_PROCESS_NAME = 4
    INT_OFFSET_PROCESS_STEP = 5
    INT_OFFSET_TASK_STATUS = 6
        
    str_static_data_provider_id = "static"
    Set col_listeners = New Collection
    
    str_id = str_module
End Sub

Public Function set_clear_data()
    obj_multi_data_provider.set_clear_data_before_close_status True
End Function

Public Function reset_clear_data()
    obj_multi_data_provider.set_clear_data_before_close_status False
End Function

Public Function load()
    obj_multi_data_provider.add_listener Me
    obj_multi_data_provider.load_data
    obj_multi_data_provider.remove_listener str_id
End Function

Public Function load_data_for_interval(obj_date_start As Date)
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    Dim str_provider_id As String
    
    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
    obj_data_provider_util.str_datetime = obj_date_start
    str_provider_id = obj_data_provider_util.str_provider_id
    On Error GoTo INFO_PROVIDER_NOT_FOUND
    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_provider_id)
    On Error GoTo 0
    
    obj_data_provider.bool_clear_before_close_status = True
    On Error GoTo INFO_FILE_NOT_EXISTS
    obj_data_provider.load_data
    On Error GoTo 0
    obj_multi_data_provider.remove_provider obj_data_provider
    Exit Function
INFO_PROVIDER_NOT_FOUND:
    Set obj_data_provider = _
        obj_multi_data_provider.create_data_provider_read_mode( _
            str_provider_id, obj_data_provider_util.retrieve_creation_file_name)
    obj_data_provider.bool_save_mode_on = True ' to be able save the file as empty one after loading is finished
    Resume Next
INFO_FILE_NOT_EXISTS:
    ' file not exists yet => remove provider
    obj_multi_data_provider.remove_provider obj_data_provider
End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_kpi_additional_data As KPIAdditionalData
    Dim obj_listener As Object

    ' fill data in a row
    Set obj_kpi_additional_data = New KPIAdditionalData
      ' pallet level
    obj_kpi_additional_data.str_id = rg_record.Offset(0, INT_OFFSET_PALLET).Value
      ' date and time
    obj_kpi_additional_data.str_date = rg_record.Offset(0, INT_OFFSET_DATE).Value
    obj_kpi_additional_data.str_time = rg_record.Offset(0, INT_OFFSET_TIME).Value
      ' place
    obj_kpi_additional_data.str_building = rg_record.Offset(0, INT_OFFSET_BUILDING).Value
      ' process
    obj_kpi_additional_data.str_process_name = rg_record.Offset(0, INT_OFFSET_PROCESS_NAME).Value
    obj_kpi_additional_data.str_process_step = rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value
      ' task status
    obj_kpi_additional_data.str_task_status = rg_record.Offset(0, INT_OFFSET_TASK_STATUS).Value
    
    For Each obj_listener In col_listeners
        obj_listener.add_data_process_record obj_kpi_additional_data
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        obj_listener.loading_data_is_about_to_start obj_provider_info
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        obj_listener.loading_data_has_finished obj_provider_info 'obj_provider_info.obj_period
    Next
End Function

Public Function save_record_dynamic(obj_kpi_additional_data As KPIAdditionalData)
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    Dim obj_mdl_processed_file_next_module As MDLProcessedFile
    Dim obj_file_record_next_module As DBProcessedFileRecord
    'Dim str_provider_id As String

    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
    obj_data_provider_util.str_datetime = obj_kpi_additional_data.str_date & " " & obj_kpi_additional_data.str_time
    'str_provider_id = obj_data_provider_util.str_provider_id
    On Error GoTo INFO_PROVIDER_NOT_FOUND
    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(obj_data_provider_util.str_provider_id)
    On Error GoTo 0

    save_record obj_kpi_additional_data, obj_data_provider.next_data
    Exit Function
INFO_PROVIDER_NOT_FOUND:
    Set obj_data_provider = _
        obj_multi_data_provider.create_data_provider_save_mode( _
            obj_data_provider_util.str_provider_id, _
            obj_data_provider_util.retrieve_creation_file_name, _
            obj_data_provider_util.retrieve_creation_tmpl_file_name)
    ' delete record in processed files of following module (e.g. kpi_result_client)
    Set obj_mdl_processed_file_next_module = New MDLProcessedFile
    obj_mdl_processed_file_next_module.str_path = STR_PROCESSED_FILE_PATH_NEXT_MODULE
    obj_mdl_processed_file_next_module.str_file_name = STR_PROCESSED_FILE_NAME_NEXT_MODULE
    obj_mdl_processed_file_next_module.open_data
    Set obj_mdl_processed_file_next_module.obj_file_record = New DBProcessedFileRecord
    obj_mdl_processed_file_next_module.obj_file_record.str_name = obj_data_provider.str_file_name
    obj_mdl_processed_file_next_module.delete_record
    obj_mdl_processed_file_next_module.close_data
        
    ' open file to start saving data
    obj_data_provider.open_data
    Resume Next
End Function

Public Function save_record_static(obj_data_process_record As DBDataProcess)
    save_record obj_data_process_record, obj_multi_data_provider.retrieve_provider(str_static_data_provider_id).next_data
End Function

Public Function save_record(obj_kpi_additional_data As KPIAdditionalData, rg_record As Range)
    ' pallet level
    rg_record.Offset(0, INT_OFFSET_PALLET).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_PALLET).Value = obj_kpi_additional_data.str_id
    ' date and time
    rg_record.Offset(0, INT_OFFSET_DATE).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_DATE).Value = obj_kpi_additional_data.str_date
    rg_record.Offset(0, INT_OFFSET_TIME).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_TIME).Value = obj_kpi_additional_data.str_time
    ' place
    rg_record.Offset(0, INT_OFFSET_BUILDING).Value = obj_kpi_additional_data.str_building
    ' process
    rg_record.Offset(0, INT_OFFSET_PROCESS_NAME).Value = obj_kpi_additional_data.str_process_name
    rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value = obj_kpi_additional_data.str_process_step
    ' task status
    rg_record.Offset(0, INT_OFFSET_TASK_STATUS).Value = obj_kpi_additional_data.str_task_status
    
End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function



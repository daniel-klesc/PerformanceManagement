VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIOnTimePallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLKPIOnTimePallet"

Public STR_DB_TYPE_FILE_EXCEL As String

Public STR_START_CHECKPOINT As String

Public INT_OFFSET_PALLET As Integer
Public INT_OFFSET_PROCESS As Integer
Public INT_OFFSET_PROCESS_VERSION As Integer
Public INT_OFFSET_PROCESS_VERSION_NAME As Integer
Public INT_OFFSET_PROCESS_STATUS As Integer
Public INT_OFFSET_MATERIAL_VENDOR As Integer
Public INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP As Integer
Public INT_OFFSET_BUILDING_SOURCE As Integer
Public INT_OFFSET_BUILDING_DESTINATION As Integer
Public INT_OFFSET_PLACE_BUILDING As Integer
Public INT_OFFSET_PLACE_FROM As Integer
Public INT_OFFSET_PLACE_TO As Integer
Public INT_OFFSET_DATE_START As Integer
Public INT_OFFSET_DATE_END As Integer
Public INT_OFFSET_DATE_LAST_CHECKPOINT As Integer
Public INT_OFFSET_INTERVAL_ID As Integer
Public INT_OFFSET_INTERVAL_START As Integer
Public INT_OFFSET_INTERVAL_END As Integer
Public INT_OFFSET_INTERVAL_SHIFT As Integer
Public INT_OFFSET_INTERVAL_SHIFT_TYPE As Integer
Public INT_OFFSET_CHECKPOINT_LIMIT As Integer

Public INT_OFFSET_KPI_STATUS As Integer


'Public INT_APP_WORK_COL_DATE_OFFSET As Integer
'Public STR_APP_WORK_FIRST_RG As String
'
'Public BYTE_STEP_STATUS As Byte

Public obj_multi_data_provider As MultiFileExcelDataProvider
Public obj_current_data_provider_info As FileExcelDataProviderInfo
Public str_static_data_provider_id As String
    ' listeners must implement:
Public str_id As String
      ' process_record obj_db_data_process: DBDataProcess
Public col_listeners As Collection
    ' listeners must implement:
      ' loading_data_is_about_to_start obj_provider_info: FileExcelDataProviderInfo
'Public col_provider_info_listeners As Collection

Public obj_md_dashboard As MDDashboard1

Private Sub Class_Initialize()
    STR_DB_TYPE_FILE_EXCEL = "file:excel"
    
    STR_START_CHECKPOINT = "START"

    INT_OFFSET_PALLET = 0
    INT_OFFSET_PROCESS = 1
    INT_OFFSET_PROCESS_VERSION = 2
    INT_OFFSET_PROCESS_VERSION_NAME = 3
    'INT_OFFSET_PROCESS_STATUS = 4
    INT_OFFSET_MATERIAL_VENDOR = 4
    INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP = 5
    INT_OFFSET_BUILDING_SOURCE = 6
    INT_OFFSET_BUILDING_DESTINATION = 7
    INT_OFFSET_PLACE_BUILDING = 8
    INT_OFFSET_PLACE_FROM = 9
    INT_OFFSET_PLACE_TO = 10
    INT_OFFSET_DATE_START = 11
    INT_OFFSET_DATE_END = 12
    INT_OFFSET_DATE_LAST_CHECKPOINT = 13
    INT_OFFSET_INTERVAL_ID = 14
    INT_OFFSET_INTERVAL_START = 15
    INT_OFFSET_INTERVAL_END = 16
    INT_OFFSET_INTERVAL_SHIFT = 17
    INT_OFFSET_INTERVAL_SHIFT_TYPE = 18
    INT_OFFSET_CHECKPOINT_LIMIT = 19

    INT_OFFSET_KPI_STATUS = 20

    
    ' app help sheet
'    INT_APP_WORK_COL_DATE_OFFSET = 0
'    STR_APP_WORK_FIRST_RG = "A2"
    
    'Set obj_multi_data_provider = New MultiFileExcelDataProvider
    str_static_data_provider_id = "static"
    Set col_listeners = New Collection
'    Set col_provider_info_listeners = New Collection
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
End Function

Public Function load_static()
    obj_multi_data_provider.add_listener Me
    obj_multi_data_provider.load_data_static
End Function

Public Function load_data_for_interval(obj_kpi_on_time_interval As KPIOnTimeInterval)
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As Object
    Dim str_provider_id As String
    
    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
    obj_data_provider_util.str_datetime = obj_kpi_on_time_interval.str_date_time
    str_provider_id = obj_data_provider_util.str_provider_id
    On Error GoTo INFO_PROVIDER_NOT_FOUND
    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_provider_id)
    On Error GoTo 0
    obj_data_provider.bool_clear_before_close_status = True
    obj_data_provider.load_data
    obj_data_provider.bool_clear_before_close_status = False
    Exit Function
INFO_PROVIDER_NOT_FOUND:
'    Set obj_data_provider = _
'        obj_multi_data_provider.create_data_provider_save_mode( _
'            str_provider_id, _
'            obj_data_provider_util.retrieve_creation_file_name, _
'            obj_data_provider_util.retrieve_creation_tmpl_file_name)
    'Resume Next
End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_kpi_on_time_pallet As KPIOnTimePallet
    Dim obj_kpi_interval As KPIOnTimeInterval
    Dim obj_interval As Interval
    Dim obj_checkpoint As KPIOnTimeLimitCheckpoint
    Dim obj_listener As Object

    ' fill data in a row
    Set obj_kpi_on_time_pallet = New KPIOnTimePallet
      ' pallet level
    obj_kpi_on_time_pallet.str_id = rg_record.Offset(0, INT_OFFSET_PALLET).Value
      ' process
    obj_kpi_on_time_pallet.str_process_id = rg_record.Offset(0, INT_OFFSET_PROCESS).Value
    obj_kpi_on_time_pallet.str_process_version_id = rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION).Value
    obj_kpi_on_time_pallet.str_process_version_name = rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_NAME).Value
      ' material
    obj_kpi_on_time_pallet.str_material_vendor = rg_record.Offset(0, INT_OFFSET_MATERIAL_VENDOR).Value
    obj_kpi_on_time_pallet.str_material_bin_storage_group = rg_record.Offset(0, INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP).Value
      ' building
    obj_kpi_on_time_pallet.str_building_source = rg_record.Offset(0, INT_OFFSET_BUILDING_SOURCE).Value
    obj_kpi_on_time_pallet.str_building_destination = rg_record.Offset(0, INT_OFFSET_BUILDING_DESTINATION).Value
      ' place
    obj_kpi_on_time_pallet.str_place_building = rg_record.Offset(0, INT_OFFSET_PLACE_BUILDING).Value
    obj_kpi_on_time_pallet.str_place_from = rg_record.Offset(0, INT_OFFSET_PLACE_FROM).Value
    obj_kpi_on_time_pallet.str_place_to = rg_record.Offset(0, INT_OFFSET_PLACE_TO).Value
      ' date and time
    obj_kpi_on_time_pallet.str_date_start = rg_record.Offset(0, INT_OFFSET_DATE_START).Value
    obj_kpi_on_time_pallet.str_date_end = rg_record.Offset(0, INT_OFFSET_DATE_END).Value
    obj_kpi_on_time_pallet.str_date_last_checkpoint = rg_record.Offset(0, INT_OFFSET_DATE_LAST_CHECKPOINT).Value
      ' interval
        ' start and end
    Set obj_interval = New Interval
    obj_interval.obj_start = rg_record.Offset(0, INT_OFFSET_INTERVAL_START).Value
    obj_interval.obj_end = rg_record.Offset(0, INT_OFFSET_INTERVAL_END).Value
    Set obj_kpi_interval = New KPIOnTimeInterval
    obj_kpi_interval.obj_interval = obj_interval
    obj_kpi_interval.str_date_time = rg_record.Offset(0, INT_OFFSET_INTERVAL_ID).Value
        ' shift
    obj_kpi_interval.str_shift = rg_record.Offset(0, INT_OFFSET_INTERVAL_SHIFT).Value
    obj_kpi_interval.str_shift_type = rg_record.Offset(0, INT_OFFSET_INTERVAL_SHIFT_TYPE).Value
    Set obj_kpi_on_time_pallet.obj_kpi_interval = obj_kpi_interval
    
      ' checkpoint
    Set obj_checkpoint = New KPIOnTimeLimitCheckpoint
    obj_checkpoint.str_next_checkpoint_id = rg_record.Offset(0, INT_OFFSET_PLACE_TO).Value
    obj_checkpoint.str_id = rg_record.Offset(0, INT_OFFSET_PLACE_FROM).Value
    obj_checkpoint.obj_limit = rg_record.Offset(0, INT_OFFSET_CHECKPOINT_LIMIT).Value
    
    If obj_checkpoint.str_id = STR_START_CHECKPOINT Then
        Set obj_kpi_on_time_pallet.obj_checkpoint = obj_checkpoint
    Else
        Set obj_kpi_on_time_pallet.obj_checkpoint = retrieve_checkpoint(obj_kpi_on_time_pallet)
    End If
    
      ' kpi status
    obj_kpi_on_time_pallet.str_kpi_status = rg_record.Offset(0, INT_OFFSET_KPI_STATUS).Value
            
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_process_record obj_kpi_on_time_pallet
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        ' obj_listener.kpi_pallet_loading_data_is_about_to_start obj_provider_info.obj_period
        obj_listener.kpi_pallet_loading_data_is_about_to_start obj_provider_info
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        'obj_listener.kpi_pallet_loading_data_has_finished obj_provider_info.obj_period
        obj_listener.kpi_pallet_loading_data_has_finished obj_provider_info
    Next
    
    Set obj_current_data_provider_info = Nothing
End Function

Public Function save_record_dynamic(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As Object
    'Dim str_provider_id As String

    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
    obj_data_provider_util.str_datetime = obj_kpi_on_time_pallet.obj_kpi_interval.str_date_time 'obj_interval.obj_start 'obj_kpi_on_time_pallet.str_date_start '& " " & obj_kpi_on_time_pallet.str_time
    'str_provider_id = obj_data_provider_util.str_provider_id
    On Error GoTo INFO_PROVIDER_NOT_FOUND
    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(obj_data_provider_util.str_provider_id)
    On Error GoTo 0

    save_record obj_kpi_on_time_pallet, obj_data_provider.next_data
    Exit Function
INFO_PROVIDER_NOT_FOUND:
    obj_data_provider_util.str_provider_status = retrieve_data_provider_completeness_status(obj_kpi_on_time_pallet.obj_kpi_interval)

    Set obj_data_provider = _
        obj_multi_data_provider.create_data_provider_save_mode( _
            obj_data_provider_util.str_provider_id, _
            obj_data_provider_util.retrieve_creation_file_name, _
            obj_data_provider_util.retrieve_creation_tmpl_file_name, _
            new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL, _
            new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_DYNAMIC)
    'obj_data_provider.obj_info.str_completeness_status = obj_data_provider_util.str_provider_status
    ' open file to start saving data
    obj_data_provider.open_data
    Resume Next
End Function

Public Function save_record_static(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim obj_data_provider As FileExcelDataProvider
    Dim rg_next_data As Range

    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_static_data_provider_id)
    On Error GoTo INFO_PROVIDER_IS_NOT_OPEN
    Set rg_next_data = obj_data_provider.next_data
    On Error GoTo 0

    save_record obj_kpi_on_time_pallet, rg_next_data
    
    Exit Function
INFO_PROVIDER_IS_NOT_OPEN:
    obj_data_provider.open_data
    Set rg_next_data = obj_data_provider.next_data
    Resume Next
End Function
'Public Function save_record_static(obj_data_process_record As DBDataProcess)
'    save_record obj_data_process_record, obj_multi_data_provider.retrieve_provider(str_static_data_provider_id).next_data
'End Function

Public Function save_record(obj_kpi_on_time_pallet As KPIOnTimePallet, rg_record As Range)
    ' pallet level
    rg_record.Offset(0, INT_OFFSET_PALLET).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_PALLET).Value = obj_kpi_on_time_pallet.str_id
     ' process
    rg_record.Offset(0, INT_OFFSET_PROCESS).Value = obj_kpi_on_time_pallet.str_process_id
    rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION).Value = obj_kpi_on_time_pallet.str_process_version_id
    rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_NAME).Value = obj_kpi_on_time_pallet.str_process_version_name
      ' material
    rg_record.Offset(0, INT_OFFSET_MATERIAL_VENDOR).Value = obj_kpi_on_time_pallet.str_material_vendor
    rg_record.Offset(0, INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP).Value = obj_kpi_on_time_pallet.str_material_bin_storage_group
      ' building
    rg_record.Offset(0, INT_OFFSET_BUILDING_SOURCE).Value = obj_kpi_on_time_pallet.str_building_source
    rg_record.Offset(0, INT_OFFSET_BUILDING_DESTINATION).Value = obj_kpi_on_time_pallet.str_building_destination
      ' place
    rg_record.Offset(0, INT_OFFSET_PLACE_BUILDING).Value = obj_kpi_on_time_pallet.str_place_building
    rg_record.Offset(0, INT_OFFSET_PLACE_FROM).Value = obj_kpi_on_time_pallet.str_place_from
    rg_record.Offset(0, INT_OFFSET_PLACE_TO).Value = obj_kpi_on_time_pallet.str_place_to
      ' date and time
    rg_record.Offset(0, INT_OFFSET_DATE_START).Value = obj_kpi_on_time_pallet.str_date_start
    rg_record.Offset(0, INT_OFFSET_DATE_END).Value = obj_kpi_on_time_pallet.str_date_end
    rg_record.Offset(0, INT_OFFSET_DATE_LAST_CHECKPOINT).Value = obj_kpi_on_time_pallet.str_date_last_checkpoint
      ' checkpoint
    rg_record.Offset(0, INT_OFFSET_CHECKPOINT_LIMIT).NumberFormat = "[$-F400]h:mm:ss AM/PM"
    rg_record.Offset(0, INT_OFFSET_CHECKPOINT_LIMIT).Value = obj_kpi_on_time_pallet.obj_checkpoint.obj_limit
      ' kpi status
    rg_record.Offset(0, INT_OFFSET_KPI_STATUS).Value = obj_kpi_on_time_pallet.str_kpi_status
    
    ' only for complete pallets
    If Not obj_kpi_on_time_pallet.obj_kpi_interval Is Nothing Then
          ' interval
            ' start, end
        rg_record.Offset(0, INT_OFFSET_INTERVAL_ID).Value = obj_kpi_on_time_pallet.obj_kpi_interval.str_id
        rg_record.Offset(0, INT_OFFSET_INTERVAL_START).Value = obj_kpi_on_time_pallet.obj_kpi_interval.obj_interval.obj_start
        rg_record.Offset(0, INT_OFFSET_INTERVAL_END).Value = obj_kpi_on_time_pallet.obj_kpi_interval.obj_interval.obj_end
            ' shift
        rg_record.Offset(0, INT_OFFSET_INTERVAL_SHIFT).Value = obj_kpi_on_time_pallet.obj_kpi_interval.str_shift
        rg_record.Offset(0, INT_OFFSET_INTERVAL_SHIFT_TYPE).Value = obj_kpi_on_time_pallet.obj_kpi_interval.str_shift_type
    End If
        
End Function

Public Function close_data()
    obj_multi_data_provider.close_providers
End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

Private Function retrieve_data_provider_completeness_status(obj_kpi_interval As KPIOnTimeInterval) As String
    If obj_kpi_interval.bool_is_finished Then
        retrieve_data_provider_completeness_status = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_STATUS_CLOSED
    Else
        retrieve_data_provider_completeness_status = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_STATUS_OPENED
    End If
End Function

Public Function retrieve_checkpoint(obj_kpi_on_time_pallet As KPIOnTimePallet) As KPIOnTimeLimitCheckpoint
    Dim bool_only_default As Boolean ' if during searching one part of limit is not found then every following part is set to default
    
    Dim str_place_checkpoint As String
    Dim obj_vendor As Vendor
    Dim obj_material_group_wh_resolver As MaterialGroupWHResolver
    Dim obj_condition_process_material As ConditionProcessMaterial
    
    Dim obj_kpi_limit_process As KPIOnTimeLimitProc
    Dim obj_kpi_limit_process_version As KPIOnTimeLimitProcVersion
    Dim obj_kpi_limit_mat_grp_wh As KPIOnTimeLimitMatGrpWH
    Dim obj_kpi_limit_source As KPIOnTimeLimitSource
    
    Dim obj_kpi As KPIOnTime
    
    Set obj_kpi = is_relevant_for_kpi(obj_kpi_on_time_pallet.str_process_version_name, obj_kpi_on_time_pallet.str_process_version_id)
    ' # implement - all retrievals must be error handled - these will lead to warning error
    Set obj_kpi_limit_process = obj_kpi.col_limits_process(obj_kpi_on_time_pallet.str_process_id)
    Set obj_kpi_limit_process_version = obj_kpi_limit_process.col_limits_version(obj_kpi_on_time_pallet.str_process_version_id)
    
    On Error GoTo WARN_NOT_FOUND_MATERIAL_GROUP_WH_RESOLVER
    Set obj_vendor = obj_md_dashboard.col_material_group_wh_resolvers(obj_kpi_on_time_pallet.str_material_vendor)
    Set obj_material_group_wh_resolver = obj_vendor.retrieve_material_group_wh_resolver(obj_kpi_on_time_pallet.str_material_bin_storage_group)
    Set obj_kpi_limit_mat_grp_wh = obj_kpi_limit_process_version.col_limits_mat_grp_wh(obj_material_group_wh_resolver.str_material_group_wh_id)
    On Error GoTo 0
    
    If Not bool_only_default Then
        On Error GoTo WARN_NOT_FOUND_CONDITION_PROCESS_MATERIAL
        
        Set obj_vendor = obj_md_dashboard.col_conditions_process_material(obj_kpi_on_time_pallet.str_material_vendor)
        Set obj_condition_process_material = obj_vendor.retrieve_condition_process_material( _
             obj_kpi_on_time_pallet.str_process_id, obj_kpi_on_time_pallet.str_process_version_id)
        Set obj_kpi_limit_source = obj_kpi_limit_mat_grp_wh.col_limits_source(obj_condition_process_material.str_condition_id)
        On Error GoTo 0
    Else
        Set obj_kpi_limit_source = obj_kpi_limit_mat_grp_wh.col_limits_source("default")
    End If
    
    On Error GoTo INFO_NOT_REGISTERED_CHECKPOINT
    Set retrieve_checkpoint = obj_kpi_limit_source.col_limits_checkpoint_place(obj_kpi_on_time_pallet.str_place_from)
    On Error GoTo 0
    
    Exit Function
WARN_NOT_FOUND_MATERIAL_GROUP_WH_RESOLVER:
    hndl_log.log db_log.TYPE_WARN, str_module, "retrieve_checkpoint", _
        "Material group wh resolver for BIN storage group:" & obj_kpi_on_time_pallet.str_material_bin_storage_group & _
        " and Vendor:" & obj_kpi_on_time_pallet.str_material_vendor & _
        " doesn't exist. Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id
    Set obj_kpi_limit_mat_grp_wh = obj_kpi_limit_process_version.col_limits_mat_grp_wh("default") ' # implement - change default string to constant
    bool_only_default = True
    Resume Next
    Exit Function
WARN_NOT_FOUND_CONDITION_PROCESS_MATERIAL:
    hndl_log.log db_log.TYPE_WARN, str_module, "retrieve_checkpoint", _
        "Condition for Process:" & obj_kpi_on_time_pallet.str_process_id & _
        ", Process version:" & obj_kpi_on_time_pallet.str_process_version_id & _
        " and Vendor:" & obj_kpi_on_time_pallet.str_material_vendor & _
        " doesn't exist. Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id
    Set obj_kpi_limit_source = obj_kpi_limit_mat_grp_wh.col_limits_source("default") ' # implement - change default string to constant
    bool_only_default = True
    Resume Next
    Exit Function
INFO_NOT_REGISTERED_CHECKPOINT:
    hndl_log.log db_log.TYPE_INFO, str_module, "retrieve_checkpoint", "Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id & ", for pallet: " & obj_kpi_on_time_pallet.str_id
End Function

Public Function is_relevant_for_kpi(str_proc_version_name As String, str_proc_version_id As String) As KPIOnTime
    Dim obj_kpi As KPIOnTime
    Dim obj_proc_ver As ProcessMasterVersion
    
    For Each obj_kpi In obj_md_dashboard.col_kpis
        For Each obj_proc_ver In obj_kpi.col_process_versions
        
            If obj_proc_ver.str_name = str_proc_version_name And obj_proc_ver.str_id = str_proc_version_id Then
                Set is_relevant_for_kpi = obj_kpi
                Exit Function
            End If
        Next
        
           
    Next
End Function

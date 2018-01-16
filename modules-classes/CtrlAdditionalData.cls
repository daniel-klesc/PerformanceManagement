VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlAdditionalData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_CLASS_NAME As String = "CtrlAdditionalData"

Private Const STR_STATUS_DONE As String = "Done"
Private Const STR_STATUS_TOTAL As String = "Total"

Private str_process_step_separator As String

Public str_id As String
Public obj_kpi As Object

Public col_result_pallets_done As Collection
Public col_result_pallets_total As Collection

Public obj_md_dashboard As MDDashboard1
Public obj_mdl_additional_data As MDLAdditionalData

' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_loading_data_is_about_to_start(obj_data_provider_info As FileExcelDataProviderInfo)
    ' # defined only because of interface
    
End Function

' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_process_record(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim str_pallet_id As String
    Dim obj_kpi_additional_data As KPIAdditionalData
    Dim obj_kpi_additional_data_done As KPIAdditionalData
    Dim bool_action_create As Boolean
    Dim obj_previous_checkpoint As KPIOnTimeLimitCheckpoint
    Dim obj_next_checkpoint As KPIOnTimeLimitCheckpoint
    
    Dim obj_process_version As ProcessMasterVersion
    Dim bool_is_interested_in_pallet As Boolean
    Dim bool_not_done_previous As Boolean

    bool_is_interested_in_pallet = False
    bool_not_done_previous = False
    ' check if KPI is interested in this process and version and checkpoint is visible
    If Not obj_kpi_on_time_pallet.obj_checkpoint.str_kpi_visibility = "Hidden" Then
        For Each obj_process_version In obj_kpi.col_process_versions
            If obj_kpi_on_time_pallet.str_process_version_name = obj_process_version.str_name Then
                bool_is_interested_in_pallet = True
                Exit For
            End If
        Next
    End If
    
    bool_action_create = False
    
    If bool_is_interested_in_pallet Then
    
        str_pallet_id = retrieve_col_result_pallet_id( _
            obj_kpi_on_time_pallet.str_id, _
            obj_kpi_on_time_pallet.str_place_from, _
            obj_kpi_on_time_pallet.str_place_to)
            
        On Error GoTo INFO_NEW_PALLET
        Set obj_kpi_additional_data = col_result_pallets_total(str_pallet_id)
        On Error GoTo 0
        
        If obj_kpi_on_time_pallet.str_kpi_status = "Closed" And Not bool_action_create Then
            Set obj_kpi_additional_data = create_pallet(obj_kpi_on_time_pallet)
            obj_kpi_additional_data.str_task_status = STR_STATUS_DONE
            col_result_pallets_done.add obj_kpi_additional_data, str_pallet_id
            
        End If
        
    End If
    
    If bool_action_create Then
               
        If Not obj_kpi_on_time_pallet.str_kpi_status = "Closed" Then
            
            obj_kpi_additional_data.str_task_status = STR_STATUS_TOTAL
            col_result_pallets_total.add obj_kpi_additional_data, str_pallet_id
            
            Set obj_previous_checkpoint = check_previous_checkpoint(obj_kpi_on_time_pallet.obj_checkpoint)
            
            If Not obj_previous_checkpoint Is Nothing Then
                Set obj_kpi_additional_data = create_pallet(obj_kpi_on_time_pallet)
                obj_kpi_additional_data.str_process_step = obj_previous_checkpoint.str_id & str_process_step_separator & obj_kpi_on_time_pallet.obj_checkpoint.str_id
                obj_kpi_additional_data.str_task_status = STR_STATUS_DONE
                col_result_pallets_done.add obj_kpi_additional_data, obj_kpi_additional_data.str_id & obj_previous_checkpoint.str_id & obj_previous_checkpoint.str_next_checkpoint_id
            End If
            
        Else
            
            obj_kpi_additional_data.str_task_status = STR_STATUS_TOTAL
            col_result_pallets_total.add obj_kpi_additional_data, str_pallet_id
            
            Set obj_previous_checkpoint = check_previous_checkpoint(obj_kpi_on_time_pallet.obj_checkpoint)
            
            On Error GoTo NOT_DONE_PREVIOUS
            If Not obj_previous_checkpoint Is Nothing Then
                Set obj_kpi_additional_data = col_result_pallets_done(obj_kpi_additional_data.str_id & obj_previous_checkpoint.str_id & obj_previous_checkpoint.str_next_checkpoint_id)
                If bool_not_done_previous Then
                    Set obj_kpi_additional_data = create_pallet(obj_kpi_on_time_pallet)
                    obj_kpi_additional_data.str_process_step = obj_previous_checkpoint.str_id & str_process_step_separator & obj_kpi_on_time_pallet.obj_checkpoint.str_id
                    obj_kpi_additional_data.str_task_status = STR_STATUS_DONE
                    col_result_pallets_done.add obj_kpi_additional_data, obj_kpi_additional_data.str_id & obj_previous_checkpoint.str_id & obj_previous_checkpoint.str_next_checkpoint_id
                End If
            End If
            On Error GoTo 0
                
            Set obj_kpi_additional_data = create_pallet(obj_kpi_on_time_pallet)
            obj_kpi_additional_data.str_task_status = STR_STATUS_DONE
            col_result_pallets_done.add obj_kpi_additional_data, str_pallet_id
        End If
        
    End If
    
    Exit Function
INFO_NEW_PALLET:
    Set obj_kpi_additional_data = create_pallet(obj_kpi_on_time_pallet)
    bool_action_create = True
    Resume Next
NOT_DONE_PREVIOUS:
    bool_not_done_previous = True
    Resume Next
End Function

' # interface method for listening MDLKPIOnTImePallet
'Public Function kpi_pallet_loading_data_has_finished(obj_period As Period)
Public Function kpi_pallet_loading_data_has_finished(obj_data_provider_info As FileExcelDataProviderInfo)
    Dim obj_kpi_additional_data As KPIAdditionalData

    For Each obj_kpi_additional_data In col_result_pallets_done
        obj_mdl_additional_data.save_record_dynamic obj_kpi_additional_data
    Next
    
    For Each obj_kpi_additional_data In col_result_pallets_total
        obj_mdl_additional_data.save_record_dynamic obj_kpi_additional_data
    Next
    
    
End Function
' # interface method for listening MDLAdditionalData
Public Function add_data_process_record(obj_kpi_add_data As KPIAdditionalData)
    Dim obj_process_version As ProcessMasterVersion
    Dim bool_is_interested_in_pallet As Boolean
    Dim arr_process_step As Variant
    
    bool_is_interested_in_pallet = False
    ' check if KPI is interested in this process and version
    For Each obj_process_version In obj_kpi.col_process_versions
        If obj_kpi_add_data.str_process_name = obj_process_version.str_name Then
            bool_is_interested_in_pallet = True
            Exit For
        End If
    Next
    If bool_is_interested_in_pallet Then
        arr_process_step = Split(obj_kpi_add_data.str_process_step, str_process_step_separator)
        
        If obj_kpi_add_data.str_task_status = STR_STATUS_DONE Then
            col_result_pallets_done.add obj_kpi_add_data, _
                retrieve_col_result_pallet_id( _
                    obj_kpi_add_data.str_id, CStr(arr_process_step(0)), CStr(arr_process_step(1)))
        ElseIf obj_kpi_add_data.str_task_status = STR_STATUS_TOTAL Then
            col_result_pallets_total.add obj_kpi_add_data, _
                retrieve_col_result_pallet_id( _
                    obj_kpi_add_data.str_id, CStr(arr_process_step(0)), CStr(arr_process_step(1)))
        End If
    End If
    
End Function

Private Function create_pallet(obj_kpi_on_time_pallet As KPIOnTimePallet) As KPIAdditionalData
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval

    Set create_pallet = New KPIAdditionalData
    create_pallet.str_id = obj_kpi_on_time_pallet.str_id
    Set obj_kpi_on_time_interval = obj_kpi_on_time_pallet.obj_kpi_interval
    create_pallet.str_date = DateValue(obj_kpi_on_time_interval.str_id)
    create_pallet.str_time = TimeValue(obj_kpi_on_time_interval.str_id)


    create_pallet.str_building = obj_kpi_on_time_pallet.str_place_building
    create_pallet.str_process_name = obj_kpi_on_time_pallet.str_process_version_name ' # implement conversion to kpi process name
    create_pallet.str_process_step = obj_kpi_on_time_pallet.str_place_from & str_process_step_separator & obj_kpi_on_time_pallet.str_place_to

End Function

Private Function retrieve_col_result_pallet_id(str_pallet As String, str_place_from As String, str_place_to As String) As String
    retrieve_col_result_pallet_id = str_pallet & str_place_from & str_place_to
End Function

Private Function check_previous_checkpoint(obj_checkpoint As KPIOnTimeLimitCheckpoint) As KPIOnTimeLimitCheckpoint
    On Error GoTo NULL_CHECKPOINT
    If obj_checkpoint.obj_previous_limit.str_kpi_visibility = "Hidden" Then
        Set check_previous_checkpoint = check_previous_checkpoint(obj_checkpoint.obj_previous_limit)
    Else
        Set check_previous_checkpoint = obj_checkpoint.obj_previous_limit
    End If
    On Error GoTo 0
    Exit Function
NULL_CHECKPOINT:
    Set check_previous_checkpoint = Nothing
End Function

Private Function check_next_checkpoint(obj_checkpoint As KPIOnTimeLimitCheckpoint) As KPIOnTimeLimitCheckpoint
    On Error GoTo NULL_CHECKPOINT
    If obj_checkpoint.obj_next_limit.str_kpi_visibility = "Hidden" Then
        Set check_next_checkpoint = check_next_checkpoint(obj_checkpoint.obj_next_limit)
    Else
        Set check_next_checkpoint = obj_checkpoint.obj_next_limit
    End If
    On Error GoTo 0
    Exit Function
NULL_CHECKPOINT:
    Set check_next_checkpoint = Nothing
End Function

Private Sub Class_Initialize()

    Set col_result_pallets_done = New Collection
    Set col_result_pallets_total = New Collection
    
    str_process_step_separator = ">"
    
End Sub

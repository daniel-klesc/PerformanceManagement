VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlKPIOnTimePallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_CLASS_NAME As String = "CtrlKPIOnTimePallet" ' # implement should be changed to KPIOnTimeResult

Public str_id As String
Public obj_kpi As Object

' data
Public obj_current_provider_info As FileExcelDataProviderInfo

Public col_db_data_process_unfinished As Collection
Public col_kpi_intervals As Collection
Public col_kpi_pallets As Collection '

'Public bool_db_data_process_unfinished_loaded As Boolean

' master data
Public obj_md_dashboard As MDDashboard1
Public obj_mdl_kpi_on_time_pallet As MDLKPIOnTimePallet ' load current update data about pallets
Public obj_mdl_kpi_on_time_pallet_unfinished As MDLKPIOnTimePallet ' load last state of active pallets

Public obj_checkpoint_type_resolver As CheckpointTypeResolver

' # method is called from MDLProcessDashboardListener
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_db_data_process As DBDataProcess

    Set obj_current_provider_info = obj_provider_info
    ' load last state of active pallets
    obj_mdl_kpi_on_time_pallet_unfinished.add_listener Me
    obj_mdl_kpi_on_time_pallet_unfinished.load_static
    create_intervals obj_provider_info.obj_period.str_start, obj_provider_info.obj_period.str_end
End Function

' # method is called from MDLProcessDashboardListener
Public Function process_record(obj_db_data_process As DBDataProcess)
    process_db_data_process_record obj_db_data_process
End Function

' # method is called from MDLProcessDashboardListener
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    save obj_provider_info
    save_unfinished obj_provider_info

    ' clean up
    Set col_kpi_intervals = New Collection
    Set col_kpi_pallets = New Collection ' # check - is it ok to clean it right here
    ' release resources
    obj_mdl_kpi_on_time_pallet.close_data
    obj_mdl_kpi_on_time_pallet_unfinished.close_data

    Set obj_current_provider_info = Nothing
End Function

Public Function create_intervals(obj_period_start As Date, obj_period_end As Date)
    Dim obj_interval As Interval ' standard interval
    Dim obj_process_interval As ProcessInterval ' master data for process
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval ' specific interval for dashboard usage
    Dim obj_kpi_on_time_interval_last As KPIOnTimeInterval ' specific interval for dashboard usage - last created
    Dim col_intervals As Collection
    
    For Each obj_process_interval In obj_kpi.obj_process.col_intervals
        Set col_intervals = fact_interval.create_intervals_hourly( _
            obj_period_start, obj_period_end, _
            obj_process_interval.int_lower_limit, obj_process_interval.int_upper_limit, obj_process_interval.int_hour_offset)
            
        For Each obj_interval In col_intervals
            Set obj_kpi_on_time_interval = New KPIOnTimeInterval
            obj_kpi_on_time_interval.obj_interval = obj_interval
            obj_kpi_on_time_interval.str_date_time = DateValue(obj_interval.obj_start) + _
                TimeSerial(Hour(obj_interval.obj_start) + obj_process_interval.int_hour_offset, obj_process_interval.int_minute, 0)
            obj_kpi_on_time_interval.bool_is_finished = CBool(obj_process_interval.int_is_complete)
            obj_kpi_on_time_interval.str_shift = hndl_master_shift.find_shift(Format(obj_interval.obj_start, "DD.MM.YYYY-h")) ' # implement through new class
            obj_kpi_on_time_interval.str_shift_type = hndl_master_shift.find_type(Hour(obj_interval.obj_start))
            col_kpi_intervals.add obj_kpi_on_time_interval, obj_kpi_on_time_interval.str_id
            
            ' create connection between kpi intervals
            If Not obj_kpi_on_time_interval_last Is Nothing Then
                Set obj_kpi_on_time_interval.obj_previous_interval = obj_kpi_on_time_interval_last
                Set obj_kpi_on_time_interval_last.obj_next_interval = obj_kpi_on_time_interval
            End If
            
            Set obj_kpi_on_time_interval_last = obj_kpi_on_time_interval
        Next
    Next
End Function

'Public Function process_db_data_process_unfinished(obj_data_process As DBDataProcess)
'       col_db_data_process_unfinished.add obj_data_process, obj_data_process.str_pallet
'End Function

' # method is called from MDLProcessDashboardListener
Public Function process_db_data_process_record(obj_data_process As DBDataProcess)
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval
    Dim obj_kpi_on_time_next_interval As KPIOnTimeInterval
    Dim obj_kpi_on_time_pallet As KPIOnTimePallet
    Dim obj_kpi_on_time_pallet_step As KPIOnTimePallet
    Dim obj_kpi_on_time_pallet_steps As Collection
    Dim obj_kpi_on_time_pallet_step_interval As KPIOnTimePallet
    Dim obj_kpi_on_time_pallet_full_live_interval As KPIOnTimePallet
    Dim str_date_end_for_status_calculation As String

    If obj_data_process.str_pallet = "357020105494051671" Then
        DoEvents
    End If

    ' check if kpi on time is interested into current obj_data_process
    If obj_kpi.obj_process.str_process_id = obj_data_process.str_creation_id Then
        On Error GoTo INFO_NEW_PALLET
        Set obj_kpi_on_time_pallet = col_kpi_pallets(obj_data_process.str_pallet)
        On Error GoTo 0
        
        If Not obj_kpi_on_time_pallet.obj_checkpoint Is Nothing Then ' check to not process pallets which already reached final checkpoint
            Set obj_kpi_on_time_pallet_steps = update_kpi_on_time_pallet(obj_kpi_on_time_pallet, obj_data_process)
        
            For Each obj_kpi_on_time_interval In col_kpi_intervals
                For Each obj_kpi_on_time_pallet_step In obj_kpi_on_time_pallet_steps
                    ' find date end for pallet
                    If obj_kpi_on_time_pallet_step.str_date_end = "" Then
                        str_date_end_for_status_calculation = obj_kpi_on_time_interval.obj_interval.obj_end
                    Else
                        str_date_end_for_status_calculation = obj_kpi_on_time_pallet_step.str_date_end
                    End If
                    
                    ' check if pallet belongs to the interval
                    If obj_kpi_on_time_interval.is_valid_interval(obj_kpi_on_time_pallet_step.str_date_last_checkpoint, str_date_end_for_status_calculation) Then
                        On Error GoTo INFO_NEW_PALLET_IN_INTERVAL
                        Set obj_kpi_on_time_pallet_step_interval = obj_kpi_on_time_interval.col_pallets_step(retrieve_pallet_id_interval(obj_kpi_on_time_pallet_step))
                        On Error GoTo 0
                        
                        ' update interval pallet
                        update_kpi_on_time_pallet_step_interval _
                            obj_kpi_on_time_interval, _
                            obj_kpi_on_time_pallet_step_interval, _
                            str_date_end_for_status_calculation
                        
                        ' update full live pallet
                        Set obj_kpi_on_time_pallet_full_live_interval = obj_kpi_on_time_pallet_step_interval.create_copy
                        
                        ' # implement - when pallet step interval is closed then remove this specific pallet step from following intervals
                        Set obj_kpi_on_time_next_interval = obj_kpi_on_time_interval.obj_next_interval
                        
                        Do While Not obj_kpi_on_time_next_interval Is Nothing
                            On Error GoTo INFO_TRY_TO_REMOVE
                            obj_kpi_on_time_next_interval.col_pallets_step.Remove retrieve_pallet_id_interval(obj_kpi_on_time_pallet_step_interval)
                            'obj_kpi_on_time_next_interval.col_pallets_full_live.Remove obj_kpi_on_time_pallet_full_live_interval.str_id ' # implement - check if this step is valid, update of full live interval pallet is done later and seperately
                            On Error GoTo 0
                            Set obj_kpi_on_time_next_interval = obj_kpi_on_time_next_interval.obj_next_interval
                        Loop
                    End If
                Next
                
                ' create
                If Not obj_kpi_on_time_pallet_full_live_interval Is Nothing Then
                    On Error GoTo INFO_NEW_PALLET_IN_FULL_LIVE_INTERVAL
                    obj_kpi_on_time_interval.col_pallets_full_live.add obj_kpi_on_time_pallet_full_live_interval, obj_kpi_on_time_pallet_full_live_interval.str_id 'retrieve_pallet_id_interval(obj_kpi_on_time_pallet_full_live_interval)
                    On Error GoTo 0
                    ' reset
                    Set obj_kpi_on_time_pallet_full_live_interval = Nothing
                End If
                
            Next
        End If
    End If
    
    Exit Function
INFO_NEW_PALLET:
    Set obj_kpi_on_time_pallet = create_kpi_on_time_pallet(obj_data_process)
    If Not obj_kpi_on_time_pallet Is Nothing Then ' it could happen if checkpoint for current place is not found or set - this behavior is required if only part of process should be measured by KPI
        Resume Next
    End If
    Exit Function
INFO_NEW_PALLET_IN_INTERVAL:
    Set obj_kpi_on_time_pallet_step_interval = obj_kpi_on_time_pallet_step.create_copy
    Set obj_kpi_on_time_pallet_step_interval.obj_kpi_interval = obj_kpi_on_time_interval
    obj_kpi_on_time_interval.col_pallets_step.add obj_kpi_on_time_pallet_step_interval, retrieve_pallet_id_interval(obj_kpi_on_time_pallet_step_interval)
    Resume Next
INFO_NEW_PALLET_IN_FULL_LIVE_INTERVAL:
    obj_kpi_on_time_interval.col_pallets_full_live.Remove obj_kpi_on_time_pallet_full_live_interval.str_id
    obj_kpi_on_time_interval.col_pallets_full_live.add obj_kpi_on_time_pallet_full_live_interval, obj_kpi_on_time_pallet_full_live_interval.str_id
    Resume Next
INFO_TRY_TO_REMOVE:
    Resume Next
End Function

Public Function create_kpi_on_time_pallet(obj_data_process As DBDataProcess) As KPIOnTimePallet
    Dim obj_checkpoint As KPIOnTimeLimitCheckpoint

    Set create_kpi_on_time_pallet = New KPIOnTimePallet
    create_kpi_on_time_pallet.str_id = obj_data_process.str_pallet
    create_kpi_on_time_pallet.str_process_id = obj_data_process.str_creation_id ' # implement - creation rename to process id
    create_kpi_on_time_pallet.str_process_version_id = obj_data_process.str_version ' # implement - version rename to process version
    create_kpi_on_time_pallet.str_process_version_name = obj_kpi.col_process_versions(obj_data_process.str_version).str_name ' # implement - version rename to process version
    create_kpi_on_time_pallet.str_material_vendor = obj_data_process.str_material_vendor
    create_kpi_on_time_pallet.str_material_bin_storage_group = obj_data_process.str_material_bin_storage_group
    create_kpi_on_time_pallet.str_building_source = _
        obj_checkpoint_type_resolver.retrieve_building(obj_data_process.str_source, obj_data_process.str_source_type)
    create_kpi_on_time_pallet.str_building_destination = _
        obj_checkpoint_type_resolver.retrieve_building(obj_data_process.str_destination, obj_data_process.str_destination_type)
        
    ' set datetime values
    create_kpi_on_time_pallet.str_date_start = obj_data_process.str_date_start
    create_kpi_on_time_pallet.str_date_last_checkpoint = obj_data_process.str_date_start
    ' set place
    create_kpi_on_time_pallet.str_place_from = bin_place_grp.get_place_grp(obj_data_process.str_bin_from)
    
    ' retrieve limit
    If obj_data_process.str_pallet = "357020105537643597" Then
        DoEvents
    End If
    
    Set obj_checkpoint = retrieve_checkpoint(create_kpi_on_time_pallet)

    If obj_checkpoint Is Nothing Then
        ' check if checkpoint exists for place_to
        create_kpi_on_time_pallet.str_place_from = bin_place_grp.get_place_grp(obj_data_process.str_bin_to)
        Set obj_checkpoint = retrieve_checkpoint(create_kpi_on_time_pallet)
        
        If obj_checkpoint Is Nothing Then
            hndl_log.log db_log.TYPE_INFO, STR_CLASS_NAME, "create_kpi_on_time_pallet", "Checkpoint not found for pallet: " & obj_data_process.str_pallet
            Set create_kpi_on_time_pallet = Nothing
        Else
            hndl_log.log db_log.TYPE_INFO, STR_CLASS_NAME, "create_kpi_on_time_pallet", "Checkpoint found in end place for pallet: " & obj_data_process.str_pallet
            update_kpi_on_time_pallet_with_new_checkpoint create_kpi_on_time_pallet, obj_checkpoint
            ' update pallet datetime data
            create_kpi_on_time_pallet.str_date_start = obj_data_process.str_date_end
            create_kpi_on_time_pallet.str_date_last_checkpoint = obj_data_process.str_date_end
            ' register pallet
            col_kpi_pallets.add create_kpi_on_time_pallet, create_kpi_on_time_pallet.str_id
        End If
    Else
        hndl_log.log db_log.TYPE_INFO, STR_CLASS_NAME, "create_kpi_on_time_pallet", "Checkpoint found in start place for pallet: " & obj_data_process.str_pallet
        update_kpi_on_time_pallet_with_new_checkpoint create_kpi_on_time_pallet, obj_checkpoint
        ' register pallet
        col_kpi_pallets.add create_kpi_on_time_pallet, create_kpi_on_time_pallet.str_id
    End If
End Function

Public Function update_kpi_on_time_pallet(obj_kpi_on_time_pallet As KPIOnTimePallet, obj_data_process As DBDataProcess) As Collection
    Dim str_place_to_new As String
    Dim obj_checkpoint As KPIOnTimeLimitCheckpoint
    Dim bool_is_place_to_tracked_by_kpi As Boolean
    Dim obj_kpi_on_time_pallet_original As KPIOnTimePallet
    
    Dim col_pallet_steps As Collection

    bool_is_place_to_tracked_by_kpi = False
    Set col_pallet_steps = New Collection

    If obj_data_process.str_pallet = "357020105430494975" Then
        DoEvents
    End If

    ' check if transaction is complete - incomplete transaction is not processed because current obj_kpi_on_time_pallet is everytime set to this last incomplete record
    If obj_data_process.str_date_end <> "" Then ' # implement - db_data_process.is_complete
        ' create copy to be able to restore pallet in original state if str_place_to is not valid checkpoint
        Set obj_kpi_on_time_pallet_original = obj_kpi_on_time_pallet.create_copy
        str_place_to_new = bin_place_grp.get_place_grp(obj_data_process.str_bin_to)

        Do
            col_pallet_steps.add obj_kpi_on_time_pallet.create_copy
            
            If Not obj_kpi_on_time_pallet.obj_checkpoint.obj_next_limit Is Nothing Then
                If obj_kpi_on_time_pallet.obj_checkpoint.obj_next_limit.str_id = str_place_to_new Then
                    bool_is_place_to_tracked_by_kpi = True
                    update_pallet_steps col_pallet_steps, obj_data_process
                    ' update pallet
                    obj_kpi_on_time_pallet.str_date_end = obj_data_process.str_date_end
                    obj_kpi_on_time_pallet.str_date_last_checkpoint = obj_data_process.str_date_end
                    Exit Do
                End If
            End If
            
            update_kpi_on_time_pallet_to_next_checkpoint obj_kpi_on_time_pallet
        Loop While Not obj_kpi_on_time_pallet.obj_checkpoint.obj_next_limit Is Nothing

        If bool_is_place_to_tracked_by_kpi Then
            ' move pallet to next checkpoint to find out if pallet has reached its final checkpoint
            update_kpi_on_time_pallet_to_next_checkpoint obj_kpi_on_time_pallet
            
            If Not obj_kpi_on_time_pallet.obj_checkpoint.obj_next_limit Is Nothing Then
                col_pallet_steps.add obj_kpi_on_time_pallet.create_copy
            End If
        Else
            ' str_place_to_new is not in tracked places (no checkpoint exists for it) => no steps are created
            Set col_pallet_steps = New Collection
            ' back to original pallet, no movement has been made
            col_kpi_pallets.Remove obj_kpi_on_time_pallet_original.str_id
            col_kpi_pallets.add obj_kpi_on_time_pallet_original, obj_kpi_on_time_pallet_original.str_id
            Set obj_kpi_on_time_pallet = obj_kpi_on_time_pallet_original.create_copy
        End If
    End If
    
    Set update_kpi_on_time_pallet = col_pallet_steps
End Function

Public Function update_kpi_on_time_pallet_with_new_checkpoint(obj_kpi_on_time_pallet As KPIOnTimePallet, obj_checkpoint As KPIOnTimeLimitCheckpoint)
    Set obj_kpi_on_time_pallet.obj_checkpoint = obj_checkpoint
    ' check if it is not last checkpoint
    If Not obj_checkpoint.obj_next_limit Is Nothing Then
        obj_kpi_on_time_pallet.str_place_to = obj_checkpoint.obj_next_limit.str_id
        obj_kpi_on_time_pallet.str_place_building = obj_checkpoint_type_resolver.retrieve_building_for_checkpoint( _
            obj_kpi_on_time_pallet.str_building_source, const_place_type.STR_IMPLICIT, _
            obj_kpi_on_time_pallet.str_building_destination, const_place_type.STR_IMPLICIT, _
            obj_kpi_on_time_pallet.obj_checkpoint.obj_next_limit)
    Else
        ' last checkpoint has no other successor, reset destination pallet data
        obj_kpi_on_time_pallet.str_place_to = ""
        obj_kpi_on_time_pallet.str_place_building = ""
    End If
End Function

Public Function update_kpi_on_time_pallet_to_next_checkpoint(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim obj_checkpoint As KPIOnTimeLimitCheckpoint
    
    ' update first part of pallet step
    obj_kpi_on_time_pallet.str_place_from = obj_kpi_on_time_pallet.str_place_to
    obj_kpi_on_time_pallet.str_date_last_checkpoint = obj_kpi_on_time_pallet.str_date_end
    
    ' reset second part of pallet step
    obj_kpi_on_time_pallet.str_date_end = ""
    obj_kpi_on_time_pallet.str_place_to = ""
    obj_kpi_on_time_pallet.str_place_building = ""
    
    ' load new checkpoint for new str_place_from
    Set obj_checkpoint = retrieve_checkpoint(obj_kpi_on_time_pallet)
    update_kpi_on_time_pallet_with_new_checkpoint obj_kpi_on_time_pallet, obj_checkpoint
End Function

Public Function update_pallet_steps(col_pallet_steps As Collection, obj_db_data_process As DBDataProcess)
    Dim obj_pallet_step As KPIOnTimePallet
    
    For Each obj_pallet_step In col_pallet_steps
        obj_pallet_step.str_date_last_checkpoint = obj_db_data_process.str_date_end
        obj_pallet_step.str_date_end = obj_db_data_process.str_date_end
    Next
End Function

Public Function update_kpi_on_time_pallet_step_interval( _
        obj_kpi_on_time_interval As KPIOnTimeInterval, _
        obj_kpi_on_time_pallet_step_interval As KPIOnTimePallet, _
        str_date_end_for_status_calculation)
            
    If CDate(str_date_end_for_status_calculation) > obj_kpi_on_time_interval.obj_interval.obj_end Then
        obj_kpi_on_time_pallet_step_interval.str_date_end = obj_kpi_on_time_interval.obj_interval.obj_end
    Else
        obj_kpi_on_time_pallet_step_interval.str_date_end = str_date_end_for_status_calculation
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
    hndl_log.log db_log.TYPE_WARN, STR_CLASS_NAME, "retrieve_checkpoint", _
        "Material group wh resolver for BIN storage group:" & obj_kpi_on_time_pallet.str_material_bin_storage_group & _
        " and Vendor:" & obj_kpi_on_time_pallet.str_material_vendor & _
        " doesn't exist. Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id
    Set obj_kpi_limit_mat_grp_wh = obj_kpi_limit_process_version.col_limits_mat_grp_wh("default") ' # implement - change default string to constant
    bool_only_default = True
    Resume Next
    Exit Function
WARN_NOT_FOUND_CONDITION_PROCESS_MATERIAL:
    hndl_log.log db_log.TYPE_WARN, STR_CLASS_NAME, "retrieve_checkpoint", _
        "Condition for Process:" & obj_kpi_on_time_pallet.str_process_id & _
        ", Process version:" & obj_kpi_on_time_pallet.str_process_version_id & _
        " and Vendor:" & obj_kpi_on_time_pallet.str_material_vendor & _
        " doesn't exist. Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id
    Set obj_kpi_limit_source = obj_kpi_limit_mat_grp_wh.col_limits_source("default") ' # implement - change default string to constant
    bool_only_default = True
    Resume Next
    Exit Function
INFO_NOT_REGISTERED_CHECKPOINT:
    hndl_log.log db_log.TYPE_INFO, STR_CLASS_NAME, "retrieve_checkpoint", "Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id & ", for pallet: " & obj_kpi_on_time_pallet.str_id
End Function

Public Function retrieve_pallet_id_interval(obj_kpi_on_time_pallet As KPIOnTimePallet)
    retrieve_pallet_id_interval = obj_kpi_on_time_pallet.str_id & obj_kpi_on_time_pallet.str_place_from
End Function

' # method is called from MDLProcessDashboardListener
Public Function save(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval
    Dim obj_kpi_on_time_pallet As KPIOnTimePallet

    For Each obj_kpi_on_time_interval In col_kpi_intervals
        ' process all pallets in interval
        For Each obj_kpi_on_time_pallet In obj_kpi_on_time_interval.col_pallets_step
            If obj_kpi_on_time_pallet.str_id = "357020105501109173" Then ' # debug - clear out
                Debug.Print "KPI interval" & obj_kpi_on_time_interval.str_id & ", Pallet data:" & obj_kpi_on_time_pallet.str_place_from & ">" & obj_kpi_on_time_pallet.str_place_to ' # debug - clear out
                DoEvents ' # debug - clear out
            End If ' # debug - clear out
        
            obj_mdl_kpi_on_time_pallet.save_record_dynamic obj_kpi_on_time_pallet
        Next
    Next
End Function

' Processing of unfinished pallets

Public Function save_unfinished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_pallet As KPIOnTimePallet

    For Each obj_pallet In col_kpi_pallets
        ' find out if pallet should be saved into unfinished
        If Not obj_pallet.obj_checkpoint.obj_next_limit Is Nothing Then
            obj_mdl_kpi_on_time_pallet_unfinished.save_record_static obj_pallet
        End If
    Next
End Function

' Listening to model with unfinished kpi pallets
' # interface method for listening MDLKPIOnTImePallet
'Public Function kpi_pallet_loading_data_is_about_to_start(obj_date_start As Date, obj_date_end As Date)
Public Function kpi_pallet_loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    obj_mdl_kpi_on_time_pallet_unfinished.obj_multi_data_provider.retrieve_provider(obj_provider_info.str_provider_id). _
        set_clear_data_before_close_status True
End Function

' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_process_record(obj_kpi_on_time_pallet As KPIOnTimePallet)
    ' register pallet
    On Error GoTo WARN_PALLET_ALREADY_EXISTS
    col_kpi_pallets.add obj_kpi_on_time_pallet, obj_kpi_on_time_pallet.str_id
    On Error GoTo 0

    Exit Function
WARN_PALLET_ALREADY_EXISTS:
    hndl_log.log db_log.TYPE_WARN, STR_CLASS_NAME, "kpi_pallet_process_record", _
        "An error occured while adding pallet: " & obj_kpi_on_time_pallet.str_id & _
        ". The original error content: " & Err.description
End Function

' # interface method for listening MDLKPIOnTImePallet
'Public Function kpi_pallet_loading_data_has_finished(obj_period As Period)
Public Function kpi_pallet_loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    ' close data for now, again data will be open for writing in save_unfinished method
    'obj_mdl_kpi_on_time_pallet_unfinished.close_data
    obj_mdl_kpi_on_time_pallet_unfinished.obj_multi_data_provider.retrieve_provider(obj_provider_info.str_provider_id). _
        set_clear_data_before_close_status False
End Function

Private Sub Class_Initialize()
    'bool_db_data_process_unfinished_loaded = False
    
    Set col_kpi_intervals = New Collection
'    Set col_result_pallets = New Collection
    Set col_kpi_pallets = New Collection
    
    Set obj_checkpoint_type_resolver = New CheckpointTypeResolver ' # implement - cross check if this is the correct place where to initialize it
End Sub

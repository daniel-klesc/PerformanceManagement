VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlAdditionalDataClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_CLASS_NAME As String = "CtrlAdditionalDataClient" ' # implement should be changed to KPIOnTimeResult

Public str_id As String
Public obj_kpi As Object

Public obj_md_dashboard As MDDashboard1
Public obj_mdl_add_data_client As MDLAdditionalDataClient

Public col_add_data_client As Collection
Public STR_ID_SEPARATOR As String

Public Function add_data_process_record(obj_add_data_pallet As KPIAdditionalData)
    Dim obj_add_data_client As KPIAdditionalDataClient
    Dim bool_reset As Boolean
    
    If is_interested_in_process(obj_add_data_pallet.str_process_name) Then
        On Error GoTo INFO_NEW_RECORD
        Set obj_add_data_client = col_add_data_client(retrieve_kpi_add_data_client_id_from_pallet(obj_add_data_pallet))
        On Error GoTo 0
        
        If Not obj_add_data_client.is_ready_for_update Then
            obj_add_data_client.reset ' for data which were loaded from result-client file contains old data
            obj_add_data_client.is_ready_for_update = True
        End If
        
        obj_add_data_client.add_result_pallet obj_add_data_pallet
    End If
    
    Exit Function
INFO_NEW_RECORD:
    Set obj_add_data_client = create_result_client(obj_add_data_pallet)
    obj_add_data_client.is_ready_for_update = True
    col_add_data_client.add obj_add_data_client, retrieve_kpi_add_data_client_id_from_pallet(obj_add_data_pallet)
    Resume Next
End Function

Public Function create_result_client(obj_result_pallet As KPIAdditionalData) As KPIAdditionalDataClient
    Dim obj_add_data_client As KPIAdditionalDataClient
    
    Set obj_add_data_client = New KPIAdditionalDataClient
    obj_add_data_client.str_date = obj_result_pallet.str_date
    obj_add_data_client.str_time = obj_result_pallet.str_time
    
    obj_add_data_client.str_shift = hndl_master_shift.find_shift(Format(obj_add_data_client.str_date & "-" & Hour(obj_add_data_client.str_time), "DD.MM.YYYY-h"))
    obj_add_data_client.str_process = obj_result_pallet.str_process_name
    obj_add_data_client.str_process_step = obj_result_pallet.str_process_step
    
    
    obj_add_data_client.str_building = obj_result_pallet.str_building
    
    obj_add_data_client.str_task_status = obj_result_pallet.str_task_status
    
    obj_add_data_client.str_datetime_shift_separator = obj_mdl_add_data_client.str_datetime_shift_separator
    obj_add_data_client.str_shift_type = hndl_master_shift.find_type(Hour(obj_add_data_client.str_time))
    'obj_result_client.str_datetime_shift = obj_result_pallet.str_date & str_datetime_shift_separator & obj_result_pallet.str_shift_type
    
    Set create_result_client = obj_add_data_client
End Function

Public Function retrieve_add_data_client_id(obj_add_data_client As KPIAdditionalDataClient) As String
    Dim str_id As String
    
    str_id = obj_add_data_client.str_date
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data_client.str_time
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data_client.str_process
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data_client.str_process_step
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data_client.str_building
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data_client.str_task_status
    
    retrieve_add_data_client_id = str_id
End Function

Public Function retrieve_kpi_add_data_client_id_from_pallet(obj_add_data As KPIAdditionalData) As String
    Dim str_id As String
    
    str_id = obj_add_data.str_date
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data.str_time
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data.str_process_name
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data.str_process_step
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data.str_building
    str_id = str_id & STR_ID_SEPARATOR & obj_add_data.str_task_status
    
    retrieve_kpi_add_data_client_id_from_pallet = str_id
End Function

Private Function is_interested_in_process(str_process_name) As Boolean
    Dim obj_process_version As ProcessMasterVersion
    
    is_interested_in_process = False
    
    For Each obj_process_version In obj_kpi.col_process_versions
        If obj_process_version.str_name = str_process_name Then
            is_interested_in_process = True
            Exit For
        End If
    Next
End Function

' # interface from MDLKPIResultDashboardListener
Public Function after_run()
    
End Function

Public Function kpi_result_loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_add_data_client As KPIAdditionalDataClient

    For Each obj_add_data_client In col_add_data_client
        obj_mdl_add_data_client.save_record_static obj_add_data_client
    Next
        
    Set col_add_data_client = New Collection
End Function

' MDLAdditionalDataClient interface
Public Function add_data_client_loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
End Function


Public Function add_data_client_process_record(obj_add_data_client As KPIAdditionalDataClient)
    If is_interested_in_process(obj_add_data_client.str_process) Then
        obj_add_data_client.is_ready_for_update = False
        col_add_data_client.add obj_add_data_client, retrieve_add_data_client_id(obj_add_data_client)
    End If
End Function


Public Function add_data_client_loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
'    Dim obj_listener As Object
'    For Each obj_listener In col_listeners
'        obj_listener.loading_data_has_finished
'    Next
End Function

Public Function kpi_result_loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
End Function

Public Function kpi_result_process_record(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
End Function


'MDLKPIOnTimeresultClient interface - not implemented
Public Function kpi_result_client_loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
End Function

Public Function kpi_result_client_process_record(obj_add_data_client As KPIOnTimeResultClient)
End Function

Public Function kpi_result_client_loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
End Function


Public Function add_mdl_kpi_on_time_result_client(obj_mdl As MDLAdditionalDataClient)
    Set obj_mdl_add_data_client = obj_mdl
    obj_mdl.add_listener Me
End Function




Private Sub Class_Initialize()
    Set col_add_data_client = New Collection
    STR_ID_SEPARATOR = "-"
    'str_datetime_shift_separator = "-"
End Sub

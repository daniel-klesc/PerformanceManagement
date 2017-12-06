VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlKPIOnTimeResultClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_CLASS_NAME As String = "CtrlKPIOnTimeResultClient" ' # implement should be changed to KPIOnTimeResult

Public str_id As String
Public obj_kpi As Object

Public obj_md_dashboard As MDDashboard1
Public obj_mdl_kpi_on_time_result_client As MDLKPIOnTimeResultClient

Public col_kpi_result_client As Collection
Public STR_ID_SEPARATOR As String
'Public str_datetime_shift_separator As String

' MDLKPIOnTImeResult interface
Public Function kpi_result_loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    ' load data from result client file
    obj_mdl_kpi_on_time_result_client.load_static
End Function

Public Function kpi_result_process_record(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
    Dim obj_result_client As KPIOnTimeResultClient
    Dim bool_reset As Boolean
    
    On Error GoTo INFO_NEW_RECORD
    Set obj_result_client = col_kpi_result_client(retrieve_kpi_result_client_id_from_pallet(obj_kpi_on_time_result_pallet))
    On Error GoTo 0
    
    If Not obj_result_client.is_ready_for_update Then
        obj_result_client.reset ' for data which were loaded from result-client file contains old data
        obj_result_client.is_ready_for_update = True
    End If
    
    obj_result_client.add_result_pallet obj_kpi_on_time_result_pallet
    
    Exit Function
INFO_NEW_RECORD:
    Set obj_result_client = create_result_client(obj_kpi_on_time_result_pallet)
    obj_result_client.is_ready_for_update = True
    col_kpi_result_client.add obj_result_client, retrieve_kpi_result_client_id_from_pallet(obj_kpi_on_time_result_pallet)
    Resume Next
End Function

Public Function kpi_result_loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_result_client As KPIOnTimeResultClient

    For Each obj_result_client In col_kpi_result_client
        obj_result_client.calculate
        obj_mdl_kpi_on_time_result_client.save_record_static obj_result_client
    Next
    
    ' clean up
    Set col_kpi_result_client = New Collection
    obj_mdl_kpi_on_time_result_client.obj_multi_data_provider.close_providers True
End Function

' MDLKPIOnTImeResultClient interface
Public Function kpi_result_client_loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    
End Function

Public Function kpi_result_client_process_record(obj_kpi_on_time_result_client As KPIOnTimeResultClient)
    obj_kpi_on_time_result_client.is_ready_for_update = False
    col_kpi_result_client.add obj_kpi_on_time_result_client, retrieve_kpi_result_client_id(obj_kpi_on_time_result_client)
End Function

Public Function kpi_result_client_loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    
End Function

Public Function add_mdl_kpi_on_time_result_client(obj_mdl As MDLKPIOnTimeResultClient)
    Set obj_mdl_kpi_on_time_result_client = obj_mdl
    obj_mdl.add_listener Me
End Function

Public Function create_result_client(obj_result_pallet As KPIOnTimeResultPallet) As KPIOnTimeResultClient
    Dim obj_result_client As KPIOnTimeResultClient
    
    Set obj_result_client = New KPIOnTimeResultClient
    obj_result_client.str_date = obj_result_pallet.str_date
    obj_result_client.str_time = obj_result_pallet.str_time
    obj_result_client.str_shift = obj_result_pallet.str_shift
    obj_result_client.str_process = obj_result_pallet.str_process_name
    obj_result_client.str_process_step = obj_result_pallet.str_process_step
    obj_result_client.str_kpi = obj_result_pallet.str_kpi_id
    obj_result_client.dbl_kpi_target = obj_kpi.dbl_target
    obj_result_client.str_building = obj_result_pallet.str_building
    obj_result_client.str_datetime_shift_separator = obj_mdl_kpi_on_time_result_client.str_datetime_shift_separator
    obj_result_client.str_shift_type = obj_result_pallet.str_shift_type
    'obj_result_client.str_datetime_shift = obj_result_pallet.str_date & str_datetime_shift_separator & obj_result_pallet.str_shift_type
    
    Set create_result_client = obj_result_client
End Function

Public Function retrieve_kpi_result_client_id(obj_result_client As KPIOnTimeResultClient) As String
    Dim str_id As String
    
    str_id = obj_result_client.str_date
    str_id = str_id & STR_ID_SEPARATOR & obj_result_client.str_time
    str_id = str_id & STR_ID_SEPARATOR & obj_result_client.str_process
    str_id = str_id & STR_ID_SEPARATOR & obj_result_client.str_process_step
    str_id = str_id & STR_ID_SEPARATOR & obj_result_client.str_kpi
    str_id = str_id & STR_ID_SEPARATOR & obj_result_client.str_building
    
    retrieve_kpi_result_client_id = str_id
End Function

Public Function retrieve_kpi_result_client_id_from_pallet(obj_result_pallet As KPIOnTimeResultPallet) As String
    Dim str_id As String
    
    str_id = obj_result_pallet.str_date
    str_id = str_id & STR_ID_SEPARATOR & obj_result_pallet.str_time
    str_id = str_id & STR_ID_SEPARATOR & obj_result_pallet.str_process_name
    str_id = str_id & STR_ID_SEPARATOR & obj_result_pallet.str_process_step
    str_id = str_id & STR_ID_SEPARATOR & obj_result_pallet.str_kpi_id
    str_id = str_id & STR_ID_SEPARATOR & obj_result_pallet.str_building
    
    retrieve_kpi_result_client_id_from_pallet = str_id
End Function

' # interface from MDLKPIResultDashboardListener
Public Function after_run()
    
End Function

Private Sub Class_Initialize()
    Set col_kpi_result_client = New Collection
    STR_ID_SEPARATOR = "-"
    'str_datetime_shift_separator = "-"
End Sub

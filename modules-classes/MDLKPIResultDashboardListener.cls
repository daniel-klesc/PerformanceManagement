VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIResultDashboardListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' # implement - all settings should be loaded through configuration file, this solution is temporary

Private Const str_module As String = "MDLKPIResultDashboardListener"

Public STR_WS_NAME As String

Public STR_TYPE_KPI As String
Public STR_SUB_TYPE_ON_TIME As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_TYPE As Integer
Public INT_OFFSET_SUB_TYPE As Integer
Public INT_OFFSET_LINK As Integer

Public obj_data_provider As FileExcelDataProvider
Public obj_md_dashboard As MDDashboard1

Public str_id As String
Public col_listeners As Collection

Public Function load()
    Set obj_data_provider = New FileExcelDataProvider
    obj_data_provider.STR_WS_NAME = STR_WS_NAME
    obj_data_provider.add_listener Me
    obj_data_provider.load_data
End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_listener As Object
    
    Set obj_listener = create(rg_record, obj_md_dashboard)
    col_listeners.add obj_listener, obj_listener.str_id
End Function

Public Function create(rg_record As Range, obj_md_dashboard As MDDashboard1) As Object
    Set create = create_type(rg_record, obj_md_dashboard)
    create.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    
    Set create.obj_md_dashboard = obj_md_dashboard
End Function

Private Function create_type(rg_record As Range, obj_md_dashboard As MDDashboard1) As Object
    Select Case rg_record.Offset(0, INT_OFFSET_TYPE).Value
        Case STR_TYPE_KPI
            Set create_type = create_kpi(rg_record, obj_md_dashboard)
    End Select
End Function

Private Function create_kpi(rg_record As Range, obj_md_dashboard As MDDashboard1) As Object
    Select Case rg_record.Offset(0, INT_OFFSET_SUB_TYPE).Value
        Case STR_SUB_TYPE_ON_TIME
            Set create_kpi = New CtrlKPIOnTimeResultClient
            On Error GoTo WARN_KPI_NOT_FOUND
            Set create_kpi.obj_kpi = obj_md_dashboard.col_kpis(rg_record.Offset(0, INT_OFFSET_LINK).Value)
            create_kpi.add_mdl_kpi_on_time_result_client create_mdl_kpi_on_time_result_client
            On Error GoTo 0
    End Select
    
    Exit Function
WARN_KPI_NOT_FOUND:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during loading data. Listener: " & rg_record.Offset(0, INT_OFFSET_ID).Value
End Function

' # interface MDLKPIOnTimeResultPallet
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_result_loading_data_is_about_to_start obj_provider_info
    Next
End Function

Public Function kpi_result_process_record(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_result_process_record obj_kpi_on_time_result_pallet
    Next
End Function

Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_result_loading_data_has_finished obj_provider_info
    Next
End Function


'' # interface method for listening MDLKPIOnTImePallet
'Public Function kpi_pallet_loading_data_is_about_to_start(obj_period As Period)
'    Dim obj_listener As Object
'    Dim obj_date_start As Date
'    Dim obj_date_end As Date
'
'    For Each obj_listener In col_listeners
'        obj_listener.kpi_pallet_loading_data_is_about_to_start obj_period.str_start, obj_period.str_end
'    Next
'End Function
'
'' # interface method for listening MDLKPIOnTImePallet
'Public Function kpi_pallet_process_record(obj_kpi_on_pallet As KPIOnTimePallet)
'    Dim obj_listener As Object
'
'    For Each obj_listener In col_listeners
'        obj_listener.kpi_pallet_process_record obj_kpi_on_pallet
'    Next
'End Function
'
'' # interface method for listening MDLKPIOnTImePallet
'Public Function kpi_pallet_loading_data_has_finished(obj_period As Period)
'    Dim obj_listener As Object
'
'    For Each obj_listener In col_listeners
'        obj_listener.kpi_pallet_loading_data_has_finished obj_period
'    Next
'End Function
'
'Public Function after_run()
''    Dim obj_listener As Object
''
''    For Each obj_listener In col_listeners
''        obj_listener.after_process
''    Next
'End Function

' THIS model is used for saving new data. Model in CtrlDashboard1 is used for loading data
Private Function create_mdl_kpi_on_time_result_client() As MDLKPIOnTimeResultClient
    ' # implement loading settings via configuration file
    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_listener As Object
    Dim str_data_provider_id As String
    
    str_data_provider_id = "final"
    
    Set create_mdl_kpi_on_time_result_client = New MDLKPIOnTimeResultClient
    create_mdl_kpi_on_time_result_client.str_static_data_provider_id = str_data_provider_id
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
          ' general
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
          ' sheet
    obj_multi_data_provider.STR_WS_NAME = "db.raw_data"
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-result-client\"
'          ' processed file
'    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
'    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_result-file_processed.xlsx" '"file_processed.xlsx"
          ' add to model
    Set create_mdl_kpi_on_time_result_client.obj_multi_data_provider = obj_multi_data_provider
    'obj_multi_data_provider.add_listener create_mdl_kpi_on_time_result_client
    
    Set obj_data_provider = New FileExcelDataProvider
    obj_data_provider.str_id = str_data_provider_id
    obj_data_provider.STR_WS_NAME = "db.raw_data"
    obj_data_provider.STR_DATA_FIRST_CELL = "A3"
    obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-result-client\"
    obj_data_provider.str_file_name = "history-result-client.xlsx"
    obj_data_provider.bool_save_mode_on = True
    obj_data_provider.add_listener create_mdl_kpi_on_time_result_client
    obj_multi_data_provider.add_provider obj_data_provider
    
'        ' data provider util
'    Set obj_data_provider_util = New FileExcelDataProviderUtil
'          ' save mode
'    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
'          ' file
'            ' tmpl
'    obj_data_provider_util.str_tmpl_path = ThisWorkbook.Path & "\data\inbound\history-result-client\tmpl\"
'    obj_data_provider_util.str_tmpl_file_name_prefix = "history-result-client"
'    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_1"
'            ' new
'    obj_data_provider_util.str_file_prefix = "history-result-client"
'    obj_data_provider_util.str_file_name_separator = "-"
'    obj_data_provider_util.str_file_appendix = ".xlsx"
'          ' add to multi data provider
'    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
End Function

Private Sub Class_Initialize()
    Set col_listeners = New Collection
    
    STR_WS_NAME = "db.md.listener"
    
    STR_TYPE_KPI = "KPI"
    STR_SUB_TYPE_ON_TIME = "ON_TIME"
    
    INT_OFFSET_ID = 0
    INT_OFFSET_TYPE = 1
    INT_OFFSET_SUB_TYPE = 2
    INT_OFFSET_LINK = 3
End Sub



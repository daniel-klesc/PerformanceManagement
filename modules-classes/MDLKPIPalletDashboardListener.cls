VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIPalletDashboardListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' # implement - all settings should be loaded through configuration file, this solution is temporary

Private Const str_module As String = "MDLKPIPalletDashboardListener"

Public STR_WS_NAME As String

Public STR_TYPE_KPI As String
Public STR_TYPE_ADD_DATA As String
Public STR_SUB_TYPE_ON_TIME As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_TYPE As Integer
Public INT_OFFSET_SUB_TYPE As Integer
Public INT_OFFSET_LINK As Integer

Public obj_data_provider As FileExcelDataProvider
Public obj_md_dashboard As MDDashboard1
Public obj_mdl_kpi_on_time_result As MDLKPIOnTimeResult
Public obj_mdl_additional_data As MDLAdditionalData

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
        Case STR_TYPE_ADD_DATA
            Set create_type = create_add_data(rg_record, obj_md_dashboard)
    End Select
End Function

Private Function create_kpi(rg_record As Range, obj_md_dashboard As MDDashboard1) As Object
    Select Case rg_record.Offset(0, INT_OFFSET_SUB_TYPE).Value
        Case STR_SUB_TYPE_ON_TIME
            If obj_mdl_kpi_on_time_result Is Nothing Then
                Set obj_mdl_kpi_on_time_result = create_mdl_kpi_on_time_result
            End If
        
            Set create_kpi = New CtrlKPIOnTimeResult
            On Error GoTo WARN_KPI_NOT_FOUND
            Set create_kpi.obj_kpi = obj_md_dashboard.col_kpis(rg_record.Offset(0, INT_OFFSET_LINK).Value)
            On Error GoTo 0
            'create_kpi.add_mdl_kpi_on_time_result obj_mdl_kpi_on_time_result
            Set create_kpi.obj_mdl_kpi_on_time_result = obj_mdl_kpi_on_time_result
    End Select
    
    Exit Function
WARN_KPI_NOT_FOUND:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during loading data. Listener: " & rg_record.Offset(0, INT_OFFSET_ID).Value
End Function

Private Function create_add_data(rg_record As Range, obj_md_dashboard As MDDashboard1) As Object
        If obj_mdl_additional_data Is Nothing Then
            Set obj_mdl_additional_data = create_mdl_additional_data
        End If
        
        Set create_add_data = New CtrlAdditionalData
        On Error GoTo WARN_KPI_NOT_FOUND
        Set create_add_data.obj_kpi = obj_md_dashboard.col_kpis(rg_record.Offset(0, INT_OFFSET_LINK).Value)
        On Error GoTo 0
        Set create_add_data.obj_mdl_additional_data = obj_mdl_additional_data
        
        Exit Function
WARN_KPI_NOT_FOUND:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during loading data. Listener: " & rg_record.Offset(0, INT_OFFSET_ID).Value
End Function

' Loading new KPI pallet data
  ' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_loading_data_is_about_to_start(obj_data_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    Dim obj_date_start As Date
    Dim obj_date_end As Date
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_loading_data_is_about_to_start obj_data_provider_info 'obj_data_provider_info.obj_period.str_start, obj_data_provider_info.obj_period.str_end
    Next
    
    ' load data from previous updates if exists for current interval (interval with ID=obj_data_provider_info.obj_period.str_start)
    obj_mdl_kpi_on_time_result.add_listener Me
    obj_mdl_kpi_on_time_result.load_data_for_interval CDate(obj_data_provider_info.obj_period.str_start)
End Function

  ' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_process_record(obj_kpi_on_pallet As KPIOnTimePallet)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_process_record obj_kpi_on_pallet
    Next
End Function

  ' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_loading_data_has_finished(obj_data_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_loading_data_has_finished obj_data_provider_info
    Next
    
    ' clean up
    obj_mdl_kpi_on_time_result.obj_multi_data_provider.close_providers
    obj_mdl_additional_data.obj_multi_data_provider.close_providers
End Function

' Loading already existing KPI result data
  ' # interface method for listening MDLKPIOnTimeResult
Public Function kpi_result_process_record(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_result_process_record obj_kpi_on_time_result_pallet
    Next
End Function

Public Function after_run()
'    Dim obj_listener As Object
'
'    For Each obj_listener In col_listeners
'        obj_listener.after_process
'    Next
End Function

Private Function create_mdl_kpi_on_time_result() As MDLKPIOnTimeResult
    ' # implement loading settings via configuration file
    Dim obj_mdl_result As MDLKPIOnTimeResult
    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    
    Dim obj_listener As Object
    
    'Set create_mdl_kpi_on_time_result = New MDLKPIOnTimeResult
    Set obj_mdl_result = New MDLKPIOnTimeResult
        ' next module - processed files
    obj_mdl_result.STR_PROCESSED_FILE_PATH_NEXT_MODULE = ThisWorkbook.Path & "\log\"
    obj_mdl_result.STR_PROCESSED_FILE_NAME_NEXT_MODULE = "history_result-file_processed.xlsx"
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
          ' general
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
          ' sheet
    obj_multi_data_provider.STR_WS_NAME = "data"
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-result\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
          ' processed file
    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_result-file_processed.xlsx" '"file_processed.xlsx"
          ' add to model
    Set obj_mdl_result.obj_multi_data_provider = obj_multi_data_provider
    obj_multi_data_provider.add_listener obj_mdl_result
        ' data provider util
    Set obj_data_provider_util = New FileExcelDataProviderUtil
          ' save mode
    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
          ' file
            ' tmpl
    obj_data_provider_util.str_tmpl_path = ThisWorkbook.Path & "\data\inbound\history-result\tmpl\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-result\tmpl\"
    obj_data_provider_util.str_tmpl_file_name_prefix = "history-result"
    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_1"
            ' new
    obj_data_provider_util.str_file_prefix = "history-result"
    obj_data_provider_util.str_file_name_separator = "-"
    obj_data_provider_util.str_file_appendix = ".xlsx"
          ' add to multi data provider
    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
    
    Set create_mdl_kpi_on_time_result = obj_mdl_result
End Function

Private Function create_mdl_additional_data() As MDLAdditionalData
    Dim obj_mdl_add_data As MDLAdditionalData
    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    
    Dim obj_listener As Object
    
    Set obj_mdl_add_data = New MDLAdditionalData
    ' next module - processed files
    obj_mdl_add_data.STR_PROCESSED_FILE_PATH_NEXT_MODULE = ThisWorkbook.Path & "\log\"
    obj_mdl_add_data.STR_PROCESSED_FILE_NAME_NEXT_MODULE = "history_add_data-file_processed.xlsx"
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
          ' general
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
          ' sheet
    obj_multi_data_provider.STR_WS_NAME = "data"
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-additional_data\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
          ' processed file
    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_additional_data-file_processed.xlsx" '"file_processed.xlsx"
          ' add to model
    Set obj_mdl_add_data.obj_multi_data_provider = obj_multi_data_provider
    obj_multi_data_provider.add_listener obj_mdl_add_data
        ' data provider util
    Set obj_data_provider_util = New FileExcelDataProviderUtil
          ' save mode
    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
          ' file
            ' tmpl
    obj_data_provider_util.str_tmpl_path = ThisWorkbook.Path & "\data\inbound\history-additional_data\tmpl\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-result\tmpl\"
    obj_data_provider_util.str_tmpl_file_name_prefix = "history-additional_data"
    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_1"
            ' new
    obj_data_provider_util.str_file_prefix = "history-additional_data"
    obj_data_provider_util.str_file_name_separator = "-"
    obj_data_provider_util.str_file_appendix = ".xlsx"
          ' add to multi data provider
    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
    
    Set create_mdl_additional_data = obj_mdl_add_data
    
    
End Function

Private Sub Class_Initialize()
    Set col_listeners = New Collection
    
    STR_WS_NAME = "db.md.listener"
    
    STR_TYPE_KPI = "KPI"
    STR_TYPE_ADD_DATA = "ADD_DATA"
    STR_SUB_TYPE_ON_TIME = "ON_TIME"
    
    INT_OFFSET_ID = 0
    INT_OFFSET_TYPE = 1
    INT_OFFSET_SUB_TYPE = 2
    INT_OFFSET_LINK = 3
End Sub


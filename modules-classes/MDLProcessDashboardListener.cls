VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLProcessDashboardListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' # implement - all settings should be loaded through configuration file, this solution is temporary

Private Const str_module As String = "MDLProcessDashboardListener"

Public STR_WS_NAME As String

Public STR_TYPE_KPI As String
Public STR_SUB_TYPE_ON_TIME As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_TYPE As Integer
Public INT_OFFSET_SUB_TYPE As Integer
Public INT_OFFSET_LINK As Integer

Public STR_DATA_PROVIDER_ID_UNFISNIHED As String

Public obj_data_provider As FileExcelDataProvider
Public obj_md_dashboard As MDDashboard1

' models for controllers
Public obj_mdl_kpi_on_time_pallet As MDLKPIOnTimePallet ' save data about pallets
Public obj_mdl_kpi_on_time_pallet_unfinished As MDLKPIOnTimePallet ' load last state of active pallets

Public str_id As String
Public col_listeners As Collection

Public Function load()
    Set obj_data_provider = New FileExcelDataProvider
    obj_data_provider.STR_WS_NAME = STR_WS_NAME
    obj_data_provider.add_listener Me
    obj_data_provider.load_data
End Function

' Loading KPI controllers
  ' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_listener As Object
    
    Set obj_listener = create(rg_record, obj_md_dashboard)
    If Not obj_listener Is Nothing Then
        col_listeners.add obj_listener, obj_listener.str_id
    End If
    'col_listeners.add obj_listener, obj_listener.str_id
End Function

Public Function create(rg_record As Range, obj_md_dashboard As MDDashboard1) As Object
    Set create = create_type(rg_record, obj_md_dashboard)
    'create.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    
    If Not create Is Nothing Then
        create.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    
        Set create.obj_md_dashboard = obj_md_dashboard
    End If
    'Set create.obj_md_dashboard = obj_md_dashboard
End Function

Private Function create_type(rg_record As Range, obj_md_dashboard As MDDashboard1) As Object
    Select Case rg_record.Offset(0, INT_OFFSET_TYPE).Value
        Case STR_TYPE_KPI
            Set create_type = create_kpi(rg_record, obj_md_dashboard)
    End Select
End Function

Private Function create_kpi(rg_record As Range, obj_md_dashboard As MDDashboard1) As Object
    Dim message As MSG

    Select Case rg_record.Offset(0, INT_OFFSET_SUB_TYPE).Value
        Case STR_SUB_TYPE_ON_TIME
            If obj_mdl_kpi_on_time_pallet Is Nothing Then
                Set obj_mdl_kpi_on_time_pallet = create_mdl_kpi_on_time_pallet()
                Set obj_mdl_kpi_on_time_pallet_unfinished = create_mdl_kpi_on_time_pallet_unfinished()
            End If
        
            Set create_kpi = New CtrlKPIOnTimePallet
            On Error GoTo WARN_KPI_NOT_FOUND
            Set create_kpi.obj_kpi = obj_md_dashboard.col_kpis(rg_record.Offset(0, INT_OFFSET_LINK).Value)
            Set create_kpi.obj_mdl_kpi_on_time_pallet = obj_mdl_kpi_on_time_pallet ' #refactoring - after change in MDLListener interface then model is handed over in process_record method
            Set create_kpi.obj_mdl_kpi_on_time_pallet_unfinished = obj_mdl_kpi_on_time_pallet_unfinished
'            ' listeners
'            Dim obj_listener As Object
'              ' data process
'            Set obj_listener = New CtrlKPIOnTimePallet
'            Set obj_listener.obj_kpi = create_kpi.obj_kpi
'            Set obj_listener.obj_mdl_kpi_on_time_pallet = create_mdl_kpi_on_time_pallet()
'            create_kpi.col_data_process_listeners.add obj_listener
'              ' kpi pallet
'            Set obj_listener = New CtrlKPIOnTimeResult
'            Set obj_listener.obj_kpi = create_kpi.obj_kpi
'            Set obj_listener.obj_mdl_kpi_on_time_result = create_mdl_kpi_on_time_result()
'            create_kpi.col_kpi_on_time_pallet_listeners.add obj_listener
            On Error GoTo 0
    End Select
    
    Exit Function
WARN_KPI_NOT_FOUND:
    Set message = New MSG
    log4VBA.warn log4VBA.DEFAULT_DESTINATION, message.source(str_module, "load_data") _
        .text("Unexpected error occurred during loading data. Listener: " & rg_record.Offset(0, INT_OFFSET_ID).Value)
End Function

' Processing of new "Process data"
  ' # interface method for listening MDLDataProcessExcel
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    Dim obj_date_start As Date
    Dim obj_date_end As Date
    
    For Each obj_listener In col_listeners
        obj_listener.loading_data_is_about_to_start obj_provider_info
    Next
    
    ' load last state of active pallets
    obj_mdl_kpi_on_time_pallet_unfinished.add_listener Me
    obj_mdl_kpi_on_time_pallet_unfinished.load_static
End Function

  ' # interface method for listening MDLDataProcessExcel
Public Function process_record(obj_db_data_process As DBDataProcess)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        On Error GoTo WARN_PROCESS_RECORD_FAILED
        obj_listener.process_record obj_db_data_process
        On Error GoTo 0
    Next
    
    Exit Function
WARN_PROCESS_RECORD_FAILED:
    Dim message As MSG
    Set message = New MSG
    log4VBA.warn log4VBA.DEFAULT_DESTINATION, message.source(str_module, "process_record").text("Unexpected error in listener:" & obj_listener.str_id & ", processing record: " & obj_db_data_process.str_pallet)
End Function

  ' # interface method for listening MDLDataProcessExcel
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
        
    For Each obj_listener In col_listeners
        obj_listener.loading_data_has_finished obj_provider_info
    Next
    
    ' release resources
    obj_mdl_kpi_on_time_pallet.close_data
    obj_mdl_kpi_on_time_pallet_unfinished.close_data
End Function


' Listening to model with unfinished kpi pallets
  ' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    obj_mdl_kpi_on_time_pallet_unfinished.obj_multi_data_provider.retrieve_provider(obj_provider_info.str_provider_id). _
        set_clear_data_before_close_status True
        
    On Error GoTo WARN_ERROR_IN_LISTENER
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_loading_data_is_about_to_start obj_provider_info
    Next
    On Error GoTo 0

    Exit Function
WARN_ERROR_IN_LISTENER:
    hndl_log.log db_log.TYPE_WARN, str_module, "kpi_pallet_loading_data_is_about_to_start", _
        "An error occured while processing listener: " & obj_listener.str_id & _
        ". The original error content: " & Err.Description
End Function

  ' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_process_record(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim obj_listener As Object

    On Error GoTo WARN_ERROR_IN_LISTENER
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_process_record obj_kpi_on_time_pallet
    Next
    On Error GoTo 0

    Exit Function
WARN_ERROR_IN_LISTENER:
    Dim log_msg As MSG
    
    Set log_msg = New MSG
'    log4VBA.warn log4VBA.DEFAULT_DESTINATION, log_msg.source(str_module, "kpi_pallet_process_record"). _
'        text("An error occured while processing listener: " & obj_listener.str_id & _
'            ". The original error content: " & Err.Description)
    
    hndl_log.log db_log.TYPE_WARN, str_module, "kpi_pallet_process_record", _
        "An error occured while processing listener: " & obj_listener.str_id & _
        ". The original error content: " & Err.Description
End Function

  ' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    On Error GoTo WARN_ERROR_IN_LISTENER
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_loading_data_has_finished obj_provider_info
    Next
    On Error GoTo 0

    obj_mdl_kpi_on_time_pallet_unfinished.obj_multi_data_provider.retrieve_provider(obj_provider_info.str_provider_id). _
        set_clear_data_before_close_status False

    Exit Function
WARN_ERROR_IN_LISTENER:
    hndl_log.log db_log.TYPE_WARN, str_module, "kpi_pallet_loading_data_is_about_to_start", _
        "An error occured while processing listener: " & obj_listener.str_id & _
        ". The original error content: " & Err.Description
End Function

'' MDLKPIOnTimePallet
'' # interface method for listening MDLKPIOnTimePallet
'Public Function kpi_pallet_loading_data_is_about_to_start(obj_data_provider_info As FileExcelDataProviderInfo)
'    Dim obj_listener As Object
'    Dim obj_date_start As Date
'    Dim obj_date_end As Date
'
'    'Debug.Print "kpi_pallet_loading_data_is_about_to_start(obj_period As Period) started"
'
'    'obj_date_start = DateValue(obj_period.str_start) + TimeValue(obj_period.str_start)
'    'obj_date_end = DateValue(obj_period.str_end) + TimeValue(obj_period.str_end)
'
'    For Each obj_listener In col_listeners
'        obj_listener.kpi_pallet_loading_data_is_about_to_start obj_data_provider_info 'obj_date_start, obj_date_end
'    Next
'End Function
'
'' # interface method for listening MDLKPIOnTimePallet
'Public Function kpi_pallet_process_record(obj_kpi_on_pallet As KPIOnTimePallet)
'    Dim obj_listener As Object
'    Debug.Print "kpi_pallet_process_record started"
'    For Each obj_listener In col_listeners
'        obj_listener.kpi_pallet_process_record obj_kpi_on_pallet
'    Next
'End Function
'
'' # interface method for listening MDLKPIOnTimePallet
'Public Function kpi_pallet_loading_data_has_finished(obj_data_provider_info As FileExcelDataProviderInfo)
'    Dim obj_listener As Object
'    Debug.Print "kpi_pallet_loading_data_has_finished started"
'    For Each obj_listener In col_listeners
'        obj_listener.save obj_data_provider_info
'    Next
'End Function

Public Function after_run()
'    Dim obj_listener As Object
'
'    For Each obj_listener In col_listeners
'        obj_listener.after_process
'    Next
End Function

Private Function create_mdl_kpi_on_time_pallet() As MDLKPIOnTimePallet
    ' # implement loading settings via configuration file
    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderStatusUtil

    Dim obj_listener As Object

    Set create_mdl_kpi_on_time_pallet = New MDLKPIOnTimePallet
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
          ' general
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
          ' sheet
    obj_multi_data_provider.STR_WS_NAME = "data"
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-kpi_pallet\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
          ' processed file
    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_kpi_pallet-file_processed.xlsx" '"file_processed.xlsx"
          ' add to model
    Set create_mdl_kpi_on_time_pallet.obj_multi_data_provider = obj_multi_data_provider
    obj_multi_data_provider.add_listener create_mdl_kpi_on_time_pallet
        ' data provider util
    Set obj_data_provider_util = New FileExcelDataProviderStatusUtil
          ' save mode
    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
          ' file
            ' tmpl
    obj_data_provider_util.str_tmpl_path = ThisWorkbook.Path & "\data\inbound\history-kpi_pallet\tmpl\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-kpi_pallet\tmpl\"
    obj_data_provider_util.str_tmpl_file_name_prefix = "history-kpi_pallet"
    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_1"
            ' new
    obj_data_provider_util.str_file_prefix = "history-kpi_pallet"
    obj_data_provider_util.str_file_name_separator = "-"
    obj_data_provider_util.str_file_appendix = ".xlsx"
            ' regex
    obj_data_provider_util.str_regex_pattern = "history-kpi_pallet-(\d{10})-(opened|closed)"
    obj_data_provider_util.int_regex_id = 0
    obj_data_provider_util.int_regex_status = 1
          ' add to multi data provider
    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
End Function

Private Function create_mdl_kpi_on_time_pallet_unfinished() As MDLKPIOnTimePallet
    Dim obj_mdl_unfinished As MDLKPIOnTimePallet
    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider As FileExcelDataProvider

    ' unfinished
    Set obj_mdl_unfinished = New MDLKPIOnTimePallet
    'obj_mdl_unfinished.BYTE_STEP_STATUS = new_ctrl_process.BYTE_STATUS_OPEN ' # implement byte_step_status should be evaluated in ctrl_pallet logic
    'obj_mdl_unfinished.str_provider_id_default = "unfinished"
    'obj_mdl_unfinished.add_listener obj_dashboard_listener
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
    'obj_multi_data_provider.STR_WS_NAME = "data" '"data.process"
    'obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_multi_data_provider.BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON = False
          ' file
    'obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-kpi_pallet\unfinished\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
            ' specific provider
    Set obj_data_provider = New FileExcelDataProvider
    obj_data_provider.str_id = "unfinished"
    obj_data_provider.STR_WS_NAME = "data"
    obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-kpi_pallet\unfinished\"
    obj_data_provider.str_file_name = "history-kpi_pallet-unfinished.xlsx"
    obj_data_provider.bool_save_mode_on = True
    obj_data_provider.add_listener obj_mdl_unfinished
    obj_multi_data_provider.add_provider obj_data_provider

    Set obj_mdl_unfinished.obj_multi_data_provider = obj_multi_data_provider
    obj_mdl_unfinished.str_static_data_provider_id = obj_data_provider.str_id
    
    Set create_mdl_kpi_on_time_pallet_unfinished = obj_mdl_unfinished
End Function

'Private Function create_mdl_kpi_on_time_result() As MDLKPIOnTimeResult
'    ' # implement loading settings via configuration file
'    Dim obj_multi_data_provider As MultiFileExcelDataProvider
'    Dim obj_data_provider_util As FileExcelDataProviderUtil
'
'    Dim obj_listener As Object
'
'    Set create_mdl_kpi_on_time_result = New MDLKPIOnTimeResult
'        ' multi data provider
'    Set obj_multi_data_provider = New MultiFileExcelDataProvider
'          ' general
'    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
'          ' sheet
'    obj_multi_data_provider.str_ws_name = "data"
'          ' file
'    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-result\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
'          ' processed file
'    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
'    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_result-file_processed.xlsx" '"file_processed.xlsx"
'          ' add to model
'    Set create_mdl_kpi_on_time_result.obj_multi_data_provider = obj_multi_data_provider
'    obj_multi_data_provider.add_listener create_mdl_kpi_on_time_result
'        ' data provider util
'    Set obj_data_provider_util = New FileExcelDataProviderUtil
'          ' save mode
'    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
'          ' file
'            ' tmpl
'    obj_data_provider_util.str_tmpl_path = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-result\tmpl\"
'    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_1"
'            ' new
'    obj_data_provider_util.str_file_prefix = "history-result"
'    obj_data_provider_util.str_file_name_separator = "-"
'    obj_data_provider_util.str_file_appendix = ".xlsx"
'          ' add to multi data provider
'    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
'End Function

Private Sub Class_Initialize()
    Set col_listeners = New Collection
    
    str_id = str_module
    
    STR_WS_NAME = "db.md.listener"
    
    STR_TYPE_KPI = "KPI"
    STR_SUB_TYPE_ON_TIME = "ON_TIME"
    
    INT_OFFSET_ID = 0
    INT_OFFSET_TYPE = 1
    INT_OFFSET_SUB_TYPE = 2
    INT_OFFSET_LINK = 3
    
    STR_DATA_PROVIDER_ID_UNFISNIHED = "unfinished"
End Sub

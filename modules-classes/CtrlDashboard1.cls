VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlDashboard1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public obj_md_dashboard As MDDashboard1

Public obj_mdl_process_dashboard_listener As MDLProcessDashboardListener
Public obj_mdl_kpi_pallet_dashboard_listener As MDLKPIPalletDashboardListener
Public obj_mdl_kpi_result_dashboard_listener As MDLKPIResultDashboardListener

Public obj_mdl_data_process As MDLDataProcessExcel
Public obj_mdl_data_process_unfinished As MDLDataProcessExcel
Public obj_mdl_data_kpi_pallet As MDLKPIOnTimePallet
Public obj_mdl_data_kpi_result As MDLKPIOnTimeResult
Public obj_current_data_process_provider_info As FileExcelDataProviderInfo

' run parameters
Public bool_run_process As Boolean
Public bool_run_process_load_unfinished As Boolean
Public bool_run_kpi_pallet As Boolean
Public bool_run_kpi_result As Boolean


Public Function before_run()
    Set obj_md_dashboard = New MDDashboard1
    obj_md_dashboard.load
    
'    Set obj_mdl_process_dashboard_listener = New MDLProcessDashboardListener
'    Set obj_mdl_process_dashboard_listener.obj_md_dashboard = obj_md_dashboard
'    obj_mdl_process_dashboard_listener.load
    
'    init_mdl_data_process obj_mdl_process_dashboard_listener
'    init_mdl_data_pallet obj_mdl_process_dashboard_listener
    'init_mdl_data_result obj_mdl_process_dashboard_listener
End Function

Public Function run()
    If bool_run_process Then
        data_process_before_run
        data_process_run
        data_process_after_run
    End If
    
    If bool_run_kpi_pallet Then
        data_kpi_pallet_before_run
        data_kpi_pallet_run
        data_kpi_pallet_after_run
    End If
    
    If bool_run_kpi_result Then
        data_kpi_result_before_run
        data_kpi_result_run
        data_kpi_result_after_run
    End If
End Function

Public Function after_run()
    ' clean up actions
'    obj_mdl_process_dashboard_listener.after_run
End Function

Public Function data_process_before_run()
    Set obj_mdl_process_dashboard_listener = New MDLProcessDashboardListener
    Set obj_mdl_process_dashboard_listener.obj_md_dashboard = obj_md_dashboard
    obj_mdl_process_dashboard_listener.load
    
    init_mdl_data_process obj_mdl_process_dashboard_listener
End Function

Public Function data_process_run()
'    If bool_run_process_load_unfinished Then
'        obj_mdl_data_process_unfinished.load_static
'    End If
    'obj_mdl_data_process_unfinished.load_static
    obj_mdl_data_process.load
End Function

Public Function data_process_after_run()
    Set new_ctrl_pallet.col_pallets = New Collection
    ' # implement - not yet decided
End Function

Public Function data_kpi_pallet_before_run()
    Set obj_mdl_kpi_pallet_dashboard_listener = New MDLKPIPalletDashboardListener
    Set obj_mdl_kpi_pallet_dashboard_listener.obj_md_dashboard = obj_md_dashboard
    obj_mdl_kpi_pallet_dashboard_listener.load
    
    init_mdl_data_pallet obj_mdl_kpi_pallet_dashboard_listener
End Function

Public Function data_kpi_pallet_run()
    obj_mdl_data_kpi_pallet.load
End Function

Public Function data_kpi_pallet_after_run()
    ' # implement - not yet decided
End Function

Public Function data_kpi_result_before_run()
    Set obj_mdl_kpi_result_dashboard_listener = New MDLKPIResultDashboardListener
    Set obj_mdl_kpi_result_dashboard_listener.obj_md_dashboard = obj_md_dashboard
    obj_mdl_kpi_result_dashboard_listener.load

    init_mdl_data_result obj_mdl_kpi_result_dashboard_listener
End Function

Public Function data_kpi_result_run()
    obj_mdl_data_kpi_result.load
End Function

Public Function data_kpi_result_after_run()
    ' # implement - not yet decided
    Dim obj_listener As Object
    
    For Each obj_listener In obj_mdl_kpi_result_dashboard_listener.col_listeners
        obj_listener.after_run
    Next
End Function

Private Function init_mdl_data_process(obj_dashboard_listener As MDLProcessDashboardListener)
    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    
    Dim obj_listener As Object

    ' # implement - all settings for initialization of obj_mdl_data_process are taken from a file
    
    ' data process model for closed records
    Set obj_mdl_data_process = New MDLDataProcessExcel
    obj_mdl_data_process.BYTE_STEP_STATUS = new_ctrl_process.BYTE_STATUS_CLOSED
    obj_mdl_data_process.add_listener obj_dashboard_listener
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
          ' general
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_multi_data_provider.BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON = True
          ' sheet
    obj_multi_data_provider.STR_WS_NAME = "data"
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-process\finished\" 'ThisWorkbook.Path & "\data\inbound\history-process\test\no-1\"  'ThisWorkbook.Path & "\data\inbound\history-process\finished\"
          ' processed file
    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_process_finished-file_processed.xlsx" '"file_processed.xlsx"
          ' add to model
    Set obj_mdl_data_process.obj_multi_data_provider = obj_multi_data_provider
        ' data provider util
    Set obj_data_provider_util = New FileExcelDataProviderUtil
          ' save mode
    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
          ' file
            ' new
    obj_data_provider_util.str_file_prefix = "history-process"
    obj_data_provider_util.str_file_name_separator = "-"
    obj_data_provider_util.str_file_appendix = ".xlsx"
          ' add to multi data provider
    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
End Function

Private Function init_mdl_data_pallet(obj_dashboard_listener As MDLKPIPalletDashboardListener)
    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderStatusUtil
    
    Dim obj_listener As Object

    ' # implement - all settings for initialization of obj_mdl_data_pallet are taken from a file
    Set obj_mdl_data_kpi_pallet = New MDLKPIOnTimePallet
    obj_mdl_data_kpi_pallet.add_listener obj_dashboard_listener
'        ' next module - processed files
'    obj_mdl_data_kpi_pallet.STR_PROCESSED_FILE_PATH_NEXT_MODULE = ThisWorkbook.Path & "\log\"
'    obj_mdl_data_kpi_pallet.STR_PROCESSED_FILE_NAME_NEXT_MODULE = "history_result-file_processed.xlsx"
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
          ' general
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_multi_data_provider.BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON = True
          ' sheet
    obj_multi_data_provider.STR_WS_NAME = "data"
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-kpi_pallet\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
          ' processed file
    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_kpi_pallet-file_processed.xlsx" '"file_processed.xlsx"
          ' add to model
    Set obj_mdl_data_kpi_pallet.obj_multi_data_provider = obj_multi_data_provider
        ' data provider util
    Set obj_data_provider_util = New FileExcelDataProviderStatusUtil
          ' save mode
    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
          ' file
            ' tmpl
    obj_data_provider_util.str_tmpl_path = ThisWorkbook.Path & "\data\inbound\history-kpi_pallet\tmpl\" ' "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-kpi_pallet\tmpl\"
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

Private Function init_mdl_data_result(obj_dashboard_listener As MDLKPIResultDashboardListener)
    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    
    Dim obj_listener As Object

    ' # implement - all settings for initialization of obj_mdl_data_kpi_result are taken from a file
    
    ' data process model for closed records
    Set obj_mdl_data_kpi_result = New MDLKPIOnTimeResult
    'obj_mdl_data_kpi_result.BYTE_STEP_STATUS = new_ctrl_process.BYTE_STATUS_CLOSED
        ' add listeners from dashboard
          ' process_provider_info obj_provider_info: FileExcelDataProviderInfo
'    obj_mdl_data_kpi_result.add_provider_info_listener obj_dashboard_listener
          ' process_record obj_db_data_process: DBDataProcess
    obj_mdl_data_kpi_result.add_listener obj_dashboard_listener
'    For Each obj_listener In obj_dashboard_listener.col_listeners
'        obj_mdl_data_kpi_result.add_listener obj_listener
'    Next

        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
          ' general
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_multi_data_provider.BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON = True
          ' sheet
    obj_multi_data_provider.STR_WS_NAME = "data"
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-result\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
          ' processed file
    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_result-file_processed.xlsx" '"file_processed.xlsx"
          ' add to model
    Set obj_mdl_data_kpi_result.obj_multi_data_provider = obj_multi_data_provider
    
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
End Function


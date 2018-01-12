Attribute VB_Name = "app"
Option Explicit

Public BOOL_CLOSE_APP As Boolean
Public int_week_beginning As Integer

Public Function init()
    Dim obj_settings As Settings
    Dim file_logger As LoggerFile
    Dim mail_logger As LoggerMail

    init = True ' if anything critical fails then init will be turned to False

    BOOL_CLOSE_APP = False
    int_week_beginning = 2
        
    bin.init
    'hndl_history.init
    'hndl_history_file_processed.init
    'hndl_log.init
    'hndl_performance.init
    'hndl_performance_output.init
    'hndl_process.init
    'hndl_proc_in_ra_vna_rack.init
    'hndl_proc_inbound_vna_in_rack.init

    ' load settings
    Set obj_settings = New Settings
    obj_settings.init
    
    On Error GoTo ERR_LOCAL_SETTING
    If Not hndl_local_setting.get_value("setting.file") = db_local_setting.STR_SETTING_FILE_DEFAULT Then
        obj_settings.str_path = hndl_local_setting.get_value("setting.file")
    End If
    On Error GoTo 0
    
    On Error GoTo ERR_OPEN_SETTINGS
    obj_settings.open_data
    On Error GoTo 0

    On Error GoTo ERR_INVALID_SETTING
    ' set new logging environment
    log4VBA.init

    Set file_logger = New LoggerFile
    file_logger.init obj_settings.Item("Performance:app.init.DEFAULT_FILE_LOGGER_NAME").Value, log4VBA.INF, log4VBA.DEFAULT_DESTINATION
    file_logger.logFilePath = ThisWorkbook.Path & obj_settings.Item("Performance:app.file_logger.logFilePath").Value
    file_logger.wsName = obj_settings.Item("Performance:file_logger.wsName").Value
    file_logger.is_same_app = True
    log4VBA.add_logger file_logger
    
    Set mail_logger = New LoggerMail
    mail_logger.init obj_settings.Item("Performance:app.init.DEFAULT_MAIL_LOGGER_NAME").Value, log4VBA.ERRO, log4VBA.DEFAULT_DESTINATION
    mail_logger.mailAddress = obj_settings.Item("Performance:mail_logger.mailAddress1").Value & ";" & obj_settings.Item("Performance:mail_logger.mailAddress2").Value
    mail_logger.subjMsgLenght = CInt(obj_settings.Item("Performance:mail_logger.subjMsgLenght").Value)
    log4VBA.add_logger mail_logger
    
'    ' hndl_log settings
'    hndl_log.str_path = obj_settings.Item("performance:file\\hndl_log.str_path").Value
'    hndl_log.str_file_name = obj_settings.Item("performance:file\\hndl_log.str_file_name.log").Value
'    ' hndl_file_processed settings
'    hndl_history_file_processed.STR_PATH_INBOUND = obj_settings.Item("performance:file\\hndl_log.str_path").Value
'    hndl_history_file_processed.str_file_name = obj_settings.Item("performance:file\\hndl_log.str_file_name.file_processed").Value
'
'    ' hndl_history
'    hndl_history.STR_PATH_INBOUND = obj_settings.Item("history pallet:file\\history.str_path_outbound").Value ' path inbound is really taken as output from application history pallet
'
'    ' hndl_performance
'    hndl_performance.STR_DAILY_WS_NAME_KPI = obj_settings.Item("performance:module\\performance.str_daily_ws_name_kpi").Value
'    hndl_performance.STR_DAILY_WS_NAME_ADDITIONAL = obj_settings.Item("performance:module\\performance.str_daily_ws_name_additional").Value
'
'    ' hndl_performance_output
'    hndl_performance_output.STR_OUTBOUND_PATH = obj_settings.Item("performance:file\\history.str_outbound_path").Value
'    hndl_performance_output.STR_OUTBOUND_FILE = obj_settings.Item("performance:file\\history.str_outbound_file").Value ' ThisWorkbook.Path & "\data\outbound\"
'    hndl_performance_output.STR_OUTBOUND_TMPL_PATH = obj_settings.Item("performance:file\\history.str_outbound_tmpl_path").Value 'ThisWorkbook.Path & "\tmpl\"
'    hndl_performance_output.str_passwd = "db_history"
'    hndl_performance_output.str_save_mode = obj_settings.Item("performance:module\\performance_output.str_save_mode").Value 'hndl_performance_output.STR_SAVE_MODE_MONTHLY
'    hndl_performance_output.STR_DAILY_WS_NAME_KPI = obj_settings.Item("performance:module\\performance_output.str_daily_ws_name_kpi").Value
'    hndl_performance_output.STR_DAILY_WS_NAME_ADDITIONAL = obj_settings.Item("performance:module\\performance_output.str_daily_ws_name_additional").Value
'    hndl_process.STR_DURATION_UNIT = "n" ' n = minutes, s = seconds
    On Error GoTo 0
    
    On Error GoTo ERR_CLOSE_SETTINGS
    obj_settings.close_data
    On Error GoTo 0
    
    Exit Function
ERR_LOCAL_SETTING:
    MsgBox "Error in local setting. Settings file not found.", vbCritical, "Application Initiation -> Loading settings"
    init = False
    Exit Function
ERR_OPEN_SETTINGS:
    MsgBox Err.Description, vbCritical, "Application Initiation -> Loading settings"
    init = False
    Exit Function
ERR_INVALID_SETTING:
    MsgBox "Invalid setting", vbCritical, "Application Initiation -> Loading settings"
    init = False
    Exit Function
ERR_CLOSE_SETTINGS:
    MsgBox "An error occured during closing settings file. Processing of history was terminated.", vbCritical, "Application Initiation -> Loading settings"
    init = False
    Exit Function
End Function

Public Function before_run()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = ""
    
    app.init
    'hndl_log.open_data
    'hndl_history_file_processed.open_data
End Function

Public Function after_run()
    Dim wb As Workbook
    
'    hndl_log.close_data
'    hndl_history_file_processed.close_data
    
    If BOOL_CLOSE_APP Then
        ThisWorkbook.Close SaveChanges:=True
        
        For Each wb In Application.Workbooks
            wb.Close SaveChanges:=False
        Next
        
        Application.Quit
    Else
        log4VBA.remove_all_loggers
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.StatusBar = False
    End If
End Function

Function run()
'    If init Then
'        BOOL_CLOSE_APP = False
'
'        before_run
'        app.init
'        after_run
'    End If
    before_run
    
    app_process.run
    app_dashboard.run
    
    after_run
End Function

Public Sub background_job()
    run
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDDashboard1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String

Public STR_MASTER_DATA_PATH As String
Public STR_MASTER_DATA_WB_NAME As String
Public STR_FILE_TYPE As String
Public BOOL_FILE_VISIBILITY As Boolean

Public STR_WS_NAME_MASTER_PROCESS As String
Public STR_WS_NAME_MASTER_PROCESS_VERSION As String
Public STR_WS_NAME_MASTER_PROCESS_INTERVAL As String

Public STR_WS_NAME_AREA_MD As String
Public STR_WS_NAME_AREA_STEP_MD As String

Public STR_WS_NAME_PLACE_MD As String

Public STR_WS_NAME_KPI As String
Public STR_WS_NAME_KPI_ON_TIME_LIMITS As String

Public STR_WS_NAME_MATERIAL_GROUP_WH_RESOLVER As String

Public STR_WS_NAME_CONDITION_PROCESS_MATERIAL As String

Public col_processes As Collection
Public col_process_versions As Collection
Public col_process_intervals As Collection
Public col_areas As Collection
Public col_area_steps As Collection
Public col_places As Collection
Public col_kpis As Collection
Public col_material_group_wh_resolvers As Collection
Public col_conditions_process_material As Collection

Public Function load()
    Dim message As MSG

    If init Then
        load_process
        load_area_steps
        load_places
        load_kpis
        load_material_group_wh_resolvers
        load_conditions_process_material
    Else
        Set message = New MSG
        log4VBA.error log4VBA.DEFAULT_DESTINATION, message.source(str_module, "init").text("Failed to load Master Data")
        Err.raise 1, "MDDashboard1.init()", "Failed to load Master Data"
    End If

End Function

Private Function load_process()
    Dim obj_mdl_process As MDLProcessMaster
    Dim obj_mdl_process_version As MDLProcessVersionMaster
    Dim obj_mdl_process_interval As MDLProcessIntervalMD

    Set obj_mdl_process = New MDLProcessMaster
    
    obj_mdl_process.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_process.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_process.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_process.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_process.single_data_provider.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS
    
    obj_mdl_process.add_listener Me
    obj_mdl_process.load_data
    
    Set obj_mdl_process_version = New MDLProcessVersionMaster
    
    obj_mdl_process_version.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_process_version.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_process_version.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_process_version.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_process_version.single_data_provider.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS_VERSION
    
    obj_mdl_process_version.add_listener Me
    obj_mdl_process_version.load_data
    
    Set obj_mdl_process_interval = New MDLProcessIntervalMD
    
    obj_mdl_process_interval.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_process_interval.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_process_interval.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_process_interval.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_process_interval.single_data_provider.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS_INTERVAL
    
    obj_mdl_process_interval.add_listener Me
    obj_mdl_process_interval.load_data
    
    Set obj_mdl_process = Nothing
    Set obj_mdl_process_version = Nothing
    Set obj_mdl_process_interval = Nothing
End Function

Private Function load_area_steps()
    Dim obj_mdl_area As MDLAreaMD
    Dim obj_mdl_area_step As MDLAreaStepMD

    Set obj_mdl_area = New MDLAreaMD
    obj_mdl_area.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_area.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_area.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_area.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_area.single_data_provider.STR_WS_NAME = STR_WS_NAME_AREA_MD '"db.md.area"
    
    obj_mdl_area.add_listener Me
    obj_mdl_area.load_data
    
    Set obj_mdl_area_step = New MDLAreaStepMD
    obj_mdl_area_step.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_area_step.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_area_step.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_area_step.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_area_step.single_data_provider.STR_WS_NAME = STR_WS_NAME_AREA_STEP_MD '"db.md.area.step"
    
    obj_mdl_area_step.add_listener Me
    obj_mdl_area_step.load_data
    
    
    Set obj_mdl_area = Nothing
    Set obj_mdl_area_step = Nothing
End Function

Private Function load_places()
    Dim obj_mdl_place As MDLPlaceMD

    Set obj_mdl_place = New MDLPlaceMD
    obj_mdl_place.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_place.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_place.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_place.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_place.single_data_provider.STR_WS_NAME = STR_WS_NAME_PLACE_MD
    
    obj_mdl_place.add_listener Me
    obj_mdl_place.load_data
    
    
    Set obj_mdl_place = Nothing
End Function

Public Function load_kpis()
    Dim obj_mdl_kpi As MDLKPIMD
    Dim obj_mdl_kpi_limit As MDLKPIOnTimeLimitsMD

    Set obj_mdl_kpi = New MDLKPIMD
    obj_mdl_kpi.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_kpi.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_kpi.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_kpi.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_kpi.single_data_provider.STR_WS_NAME = STR_WS_NAME_KPI
    
    obj_mdl_kpi.add_listener Me
    obj_mdl_kpi.load_data
    
    Set obj_mdl_kpi_limit = New MDLKPIOnTimeLimitsMD
    obj_mdl_kpi_limit.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_kpi_limit.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_kpi_limit.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_kpi_limit.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_kpi_limit.single_data_provider.STR_WS_NAME = STR_WS_NAME_KPI_ON_TIME_LIMITS
    
    obj_mdl_kpi_limit.add_listener Me
    obj_mdl_kpi_limit.load_data
    
    'Set col_kpis = obj_mdl_kpi.col_kpis
End Function

Public Function load_material_group_wh_resolvers()
    Dim obj_mdl_material_group_wh_resolver As MDLMaterialGroupWHResolver
    
    Set obj_mdl_material_group_wh_resolver = New MDLMaterialGroupWHResolver
    obj_mdl_material_group_wh_resolver.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_material_group_wh_resolver.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_material_group_wh_resolver.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_material_group_wh_resolver.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_material_group_wh_resolver.single_data_provider.STR_WS_NAME = STR_WS_NAME_MATERIAL_GROUP_WH_RESOLVER
    
    obj_mdl_material_group_wh_resolver.add_listener Me
    obj_mdl_material_group_wh_resolver.load_data
End Function

Public Function load_conditions_process_material()
    Dim obj_mdl_condition_process_material As MDLConditionProcessMaterial
    
    Set obj_mdl_condition_process_material = New MDLConditionProcessMaterial
    obj_mdl_condition_process_material.single_data_provider.str_path = STR_MASTER_DATA_PATH
    obj_mdl_condition_process_material.single_data_provider.str_file_name = STR_MASTER_DATA_WB_NAME
    obj_mdl_condition_process_material.single_data_provider.BOOL_EXTERNAL_DATA_FILE_VISIBILITY = BOOL_FILE_VISIBILITY
    obj_mdl_condition_process_material.single_data_provider.str_source_type = STR_FILE_TYPE
    obj_mdl_condition_process_material.single_data_provider.STR_WS_NAME = STR_WS_NAME_CONDITION_PROCESS_MATERIAL
    
    obj_mdl_condition_process_material.add_listener Me
    obj_mdl_condition_process_material.load_data
    'Set col_conditions_process_material = obj_mdl_condition_process_material.col_vendors
End Function

Public Function init() As Boolean
    Dim obj_settings As Settings
    init = True
    
    ' load settings
    Set obj_settings = New Settings
    obj_settings.init
    
    On Error GoTo ERR_LOCAL_SETTING
    If Not hndl_local_setting.get_value("MDsetting.file") = db_local_setting.STR_SETTING_FILE_DEFAULT Then
        obj_settings.str_path = hndl_local_setting.get_value("MDsetting.file")
    Else
        obj_settings.str_path = ThisWorkbook.Path & "\MDsettings.xlsx"
    End If
    On Error GoTo 0
    
    On Error GoTo ERR_OPEN_SETTINGS
    obj_settings.open_data
    On Error GoTo 0
    
    STR_FILE_TYPE = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    BOOL_FILE_VISIBILITY = new_const_excel_data_provider.BOOL_VISIBILITY_MODE_HIDDEN
    

    On Error GoTo ERR_INVALID_SETTING
    
    STR_MASTER_DATA_PATH = obj_settings.Item("Performance:MDDashboard1.STR_MD_PATH").Value
    STR_MASTER_DATA_WB_NAME = obj_settings.Item("Performance:MDDashboard1.STR_MD_WB_NAME").Value

    
    STR_WS_NAME_MASTER_PROCESS = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_MASTER_PROCESS").Value
    STR_WS_NAME_MASTER_PROCESS_VERSION = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_MASTER_PROCESS_VERSION").Value
    STR_WS_NAME_MASTER_PROCESS_INTERVAL = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_MASTER_PROCESS_INTERVAL").Value
    
    STR_WS_NAME_AREA_MD = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_AREA_MD").Value
    STR_WS_NAME_AREA_STEP_MD = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_AREA_STEP_MD").Value
    
    STR_WS_NAME_PLACE_MD = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_PLACE_MD").Value
    
    STR_WS_NAME_KPI = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_KPI").Value
    STR_WS_NAME_KPI_ON_TIME_LIMITS = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_KPI_ON_TIME_LIMITS").Value
    
    STR_WS_NAME_MATERIAL_GROUP_WH_RESOLVER = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_MATERIAL_GROUP_WH_RESOLVER").Value
    
    STR_WS_NAME_CONDITION_PROCESS_MATERIAL = obj_settings.Item("Performance:MDDashboard1.STR_WS_NAME_CONDITION_PROCESS_MATERIAL").Value
    
    
    On Error GoTo 0
    
    On Error GoTo ERR_CLOSE_SETTINGS
    obj_settings.close_data
    On Error GoTo 0
    
    Exit Function
ERR_LOCAL_SETTING:
    init = False
    Exit Function
ERR_OPEN_SETTINGS:
    init = False
    Exit Function
ERR_INVALID_SETTING:
    init = False
    Exit Function
ERR_CLOSE_SETTINGS:
    init = False
    Exit Function
    
End Function

Private Sub Class_Initialize()
    
    str_module = "MDDashboard1"
    
    Set col_areas = New Collection
    Set col_area_steps = New Collection
    Set col_places = New Collection
    Set col_kpis = New Collection
    Set col_material_group_wh_resolvers = New Collection
    Set col_conditions_process_material = New Collection
    Set col_processes = New Collection
    Set col_process_versions = New Collection
    Set col_process_intervals = New Collection
       
    
End Sub

Public Function get_collection(mdl_id As String) As Collection
    Select Case mdl_id
        Case new_const_mdl_id.STR_MDL_AREA:
            Set get_collection = col_areas
        
        Case new_const_mdl_id.STR_MDL_PLACE:
            Set get_collection = col_places
            
        Case new_const_mdl_id.STR_MDL_AREA_STEP:
            Set get_collection = col_area_steps
        
        Case new_const_mdl_id.STR_MDL_KPI, new_const_mdl_id.STR_MDL_KPI_ONTIME_LIMITS:
            Set get_collection = col_kpis
        
        Case new_const_mdl_id.STR_MDL_PROCESS_MASTER:
            Set get_collection = col_processes
        
        Case new_const_mdl_id.STR_MDL_MATERIAL_GROUP_WH_RESOLVER:
            Set get_collection = col_material_group_wh_resolvers
            
        Case new_const_mdl_id.STR_MDL_CONDITION_PROCESS_MATERIAL:
            Set get_collection = col_conditions_process_material
        
        Case new_const_mdl_id.STR_MDL_PROCESS_INTERVAL:
            Set get_collection = col_process_intervals
        
        Case new_const_mdl_id.STR_MDL_PROCESS_VERSION_MASTER:
            Set get_collection = col_process_versions
        
    End Select
            
End Function

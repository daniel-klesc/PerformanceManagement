VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionInboundGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public STR_VERSION_SINGLE As String
Public STR_VERSION_CROSS_B_C As String
Public STR_VERSION_CROSS_A_BC As String

Public STR_STORAGE_TYPE_VNA As String
Public STR_STORAGE_TYPE_TA As String
Public STR_STORAGE_TYPE_PROD_HALL As String
Public STR_STORAGE_TYPE_JIRNY As String

Public STR_MODULE_EXT As String
Public STR_MODULE_PROD As String

Public str_version_id As String
Public str_source As String
Public str_source_type As String
Public str_destination As String
Public str_destination_type As String


Public Function init(obj_record As DBHistoryRecord)
    Dim str_building_from As String
    Dim str_building_to As String
    Dim str_module As String
    Dim str_storage_type As String
    Dim str_sub_version As String
    
    ' from
    str_source = obj_record.str_bin_to
    str_source_type = const_place_type.str_bin
    str_building_from = bin.get_building(obj_record.str_bin_to)
    
    'to
    str_destination = obj_record.str_storage_group_material 'obj_record.str_machine_transport_ref
    str_destination_type = const_place_type.str_bin_stor_grp
    str_building_to = bin_stor_grp.get_building(obj_record.str_storage_group_material)
    If str_building_to = bin_stor_grp.STR_BUILDING_GENERAL Then
        str_building_to = str_building_from
    End If
    
    If str_building_from = "" And obj_record.str_transaction_type_started = db_transaction_type.STR_TRANSACTION_TYPE_PO_GR Then
        str_building_from = str_building_to
    End If
    
'    str_module = wc.get_module(obj_record.str_wc_ship_to)
    str_module = resolve_module(obj_record)
    
    If bin_stor_grp.is_production_hall(obj_record.str_storage_group_material) Then
        str_storage_type = STR_STORAGE_TYPE_PROD_HALL
    ElseIf bin_stor_grp.is_ta(obj_record.str_storage_group_material) Then
        str_storage_type = STR_STORAGE_TYPE_TA
    ElseIf bin_stor_grp.is_outbound(obj_record.str_storage_group_material) Then
        str_storage_type = STR_STORAGE_TYPE_JIRNY
    Else
        str_storage_type = STR_STORAGE_TYPE_VNA
    End If
    
    Select Case str_building_from & str_building_to
        Case "AA", "BB", "CC"
            str_sub_version = STR_VERSION_SINGLE 'new_ctrl_process_master_version.STR_VERSION_SINGLE
        Case "BC", "CB"
            str_sub_version = STR_VERSION_CROSS_B_C  'new_ctrl_process_master_version.STR_VERSION_SINGLE
        Case "AB", "BA", "AC", "CA"
            str_sub_version = STR_VERSION_CROSS_A_BC
    End Select
    
    str_version_id = str_module & "-" & str_storage_type & "-" & str_sub_version
End Function

Public Function resolve_module(obj_record As DBHistoryRecord) As String
    
    Select Case obj_record.str_transaction_type_started
        Case db_transaction_type.STR_TRANSACTION_TYPE_PO_GR, _
                db_transaction_type.STR_TRANSACTION_TYPE_CONS_OWN_STOCK
            resolve_module = STR_MODULE_EXT
        Case db_transaction_type.STR_TRANSACTION_TYPE_PROD_ORD_GR
            resolve_module = wc.get_module(obj_record.str_wc_ship_to)
    End Select
End Function

Private Sub Class_Initialize()
    STR_VERSION_SINGLE = "1"
    STR_VERSION_CROSS_B_C = "2"
    STR_VERSION_CROSS_A_BC = "3"
    
    STR_STORAGE_TYPE_VNA = "VNA"
    STR_STORAGE_TYPE_PROD_HALL = "PROD_HALL"
    STR_STORAGE_TYPE_TA = "TA"
    STR_STORAGE_TYPE_JIRNY = "JIRNY"
    
    STR_MODULE_EXT = "EXT"
End Sub


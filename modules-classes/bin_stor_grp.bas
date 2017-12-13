Attribute VB_Name = "bin_stor_grp"
Option Explicit

' #dependency - bin module

Public Const STR_OUTBOUND_JIRNY As String = "JIRNY"
Public Const STR_INBOUND_HBW As String = "HBW"
Public Const STR_INBOUND_PROCESSING_PAINT As String = "PAINT"

Public Const STR_BUILDING_A_PREFIX As String = "V"
Public Const STR_BUILDING_B_PREFIX As String = "B"
Public Const STR_BUILDING_C_PREFIX As String = "C"

Public Const STR_BUILDING_GENERAL As String = "GENERAL"

Public Function is_outbound(str_bin_stor_grp As String)
    is_outbound = False

    If str_bin_stor_grp = STR_OUTBOUND_JIRNY Then
        is_outbound = True
    End If
End Function

Public Function is_hbw(str_bin_stor_grp As String)
    is_hbw = False

    If str_bin_stor_grp = STR_INBOUND_HBW Then
        is_hbw = True
    End If
End Function

Public Function is_processing(str_bin_stor_grp As String)
    is_processing = False

    If str_bin_stor_grp = STR_INBOUND_PROCESSING_PAINT Then
        is_processing = True
    End If
End Function

Public Function get_building(str_bin_stor_grp As String) As String
    Dim str_building As String
    
    str_building = Left(str_bin_stor_grp, 1)
    
    Select Case str_building
        Case STR_BUILDING_A_PREFIX
            get_building = bin.STR_BUILDING_A
        Case STR_BUILDING_B_PREFIX
            get_building = bin.STR_BUILDING_B
        Case STR_BUILDING_C_PREFIX
            get_building = bin.STR_BUILDING_C
        Case Else
            get_building = STR_BUILDING_GENERAL
    End Select
End Function

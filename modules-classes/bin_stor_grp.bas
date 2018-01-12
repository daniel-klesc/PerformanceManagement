Attribute VB_Name = "bin_stor_grp"
Option Explicit

' #dependency - bin module

Public Const STR_OUTBOUND_JIRNY As String = "JIRNY"
Public Const STR_INBOUND_HBW As String = "HBW"
Public Const STR_INBOUND_PROCESSING_PAINT As String = "PAINT"

Public Const STR_BUILDING_A_PREFIX As String = "V"
Public Const STR_BUILDING_B_PREFIX As String = "B"
Public Const STR_BUILDING_C_PREFIX As String = "C"

Public Const STR_HALL_C2 As String = "HALA C2"

Public Const INT_TA_PREFIX_LEN As Integer = 2
Public Const STR_TA_PREFIX As String = "TA"
Public Const STR_TA_1 As String = "TA1"
Public Const STR_TA_2 As String = "TA2"
Public Const STR_TA_3 As String = "TA3"
Public Const STR_TA_4 As String = "TA4"
Public Const STR_TA_5 As String = "TA5"
Public Const STR_TA_6 As String = "TA6"
Public Const STR_TA_7 As String = "TA7"
Public Const STR_TA_8 As String = "TA8"

Public Const STR_BUILDING_GENERAL As String = "GENERAL"

Public Function is_outbound(str_bin_stor_grp As String) As Boolean
    is_outbound = False

    If str_bin_stor_grp = STR_OUTBOUND_JIRNY Then
        is_outbound = True
    End If
End Function

Public Function is_hbw(str_bin_stor_grp As String) As Boolean
    is_hbw = False

    If str_bin_stor_grp = STR_INBOUND_HBW Then
        is_hbw = True
    End If
End Function

Public Function is_processing(str_bin_stor_grp As String) As Boolean
    is_processing = False

    If str_bin_stor_grp = STR_INBOUND_PROCESSING_PAINT Then
        is_processing = True
    End If
End Function

Public Function is_production_hall(str_bin_stor_grp As String) As Boolean
    is_production_hall = False

    If str_bin_stor_grp = STR_HALL_C2 Then
        is_production_hall = True
    End If
End Function

Public Function is_ta(str_bin_stor_grp As String) As Boolean
    is_ta = False

    If Left(str_bin_stor_grp, INT_TA_PREFIX_LEN) = STR_TA_PREFIX Then
        is_ta = True
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
            If is_ta(str_bin_stor_grp) Then
                Select Case str_bin_stor_grp
                    Case STR_TA_1, STR_TA_2
                        get_building = bin.STR_BUILDING_A
                    Case STR_TA_3, STR_TA_4, STR_TA_5
                        get_building = bin.STR_BUILDING_B
                    Case STR_TA_6, STR_TA_7, STR_TA_8
                        get_building = bin.STR_BUILDING_C
                End Select
            ElseIf str_bin_stor_grp = STR_HALL_C2 Then ' not a general solution - quick fix
                get_building = bin.STR_BUILDING_C
            Else
                get_building = STR_BUILDING_GENERAL
            End If
    End Select
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimeResultPallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String

Public str_date As String
Public str_time As String

Public str_shift As String
Public str_shift_type As String

Public str_building As String
Public str_process_name As String
Public str_process_step As String

Public str_kpi_id As String
Public str_kpi_name As String
Public str_kpi_status As String

Public Function toString() As String
    toString = "str_id" & const_separator.SEPARATOR_VALUE & str_id & const_separator.SEPARATOR_CLASS_ELEMENT & "str_date" & const_separator.SEPARATOR_VALUE & str_date & _
    const_separator.SEPARATOR_CLASS_ELEMENT & "str_time" & const_separator.SEPARATOR_VALUE & str_time & const_separator.SEPARATOR_CLASS_ELEMENT & "str_shift" & _
    const_separator.SEPARATOR_VALUE & str_shift & const_separator.SEPARATOR_CLASS_ELEMENT & "str_shift_type" & const_separator.SEPARATOR_VALUE & str_shift_type & _
    const_separator.SEPARATOR_CLASS_ELEMENT & "str_building" & const_separator.SEPARATOR_VALUE & str_building & const_separator.SEPARATOR_CLASS_ELEMENT & _
    "str_process_name" & const_separator.SEPARATOR_VALUE & str_process_name & const_separator.SEPARATOR_CLASS_ELEMENT & "str_process_step" & const_separator.SEPARATOR_VALUE & _
    str_process_step & const_separator.SEPARATOR_CLASS_ELEMENT & "str_kpi_id" & const_separator.SEPARATOR_VALUE & str_kpi_id & const_separator.SEPARATOR_CLASS_ELEMENT & _
    "str_kpi_name" & const_separator.SEPARATOR_VALUE & str_kpi_name & const_separator.SEPARATOR_CLASS_ELEMENT & "str_kpi_status" & const_separator.SEPARATOR_VALUE & str_kpi_status & _
    const_separator.SEPARATOR_CLASS_ELEMENT
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimeLimitCheckpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String
Public str_type As String
Public str_next_checkpoint_id As String
Public obj_limit As Date
Public str_kpi_visibility As String

Public obj_previous_limit As KPIOnTimeLimitCheckpoint
Public obj_next_limit As KPIOnTimeLimitCheckpoint

Public Function toString() As String
    toString = "str_id" & const_separator.SEPARATOR_VALUE & str_id
End Function



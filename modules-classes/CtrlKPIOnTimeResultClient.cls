VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlKPIOnTimeResultClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_CLASS_NAME As String = "CtrlKPIOnTimeResultClient" ' # implement should be changed to KPIOnTimeResult

Public str_id As String
Public obj_kpi As Object

Public obj_md_dashboard As MDDashboard1
Public obj_mdl_kpi_on_time_result_client As MDLKPIOnTimeResultClient

Public Function add_mdl_kpi_on_time_result_client(obj_mdl As MDLKPIOnTimeResultClient)
    Set obj_mdl_kpi_on_time_result_client = obj_mdl
    obj_mdl.add_listener Me
End Function

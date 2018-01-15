VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListenerHistoryToProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function before_process_history_record(obj_data_provider_info As FileExcelDataProviderInfo)
    Set new_ctrl_pallet.obj_data_provider_info = obj_data_provider_info
    new_ctrl_pallet.obj_mdl_unfinished.set_clear_data
    new_ctrl_pallet.obj_mdl_unfinished.load_static
End Function

Public Function process_history_record(obj_record As DBHistoryRecord)
    new_ctrl_pallet.process_record obj_record
End Function

Public Function after_process_history_record(obj_data_provider_info As FileExcelDataProviderInfo)
    new_ctrl_pallet.obj_mdl_unfinished.reset_clear_data
    new_ctrl_pallet.save_open_pallets
End Function

Attribute VB_Name = "test_bin_place_grp"
Option Explicit

Public Function setup()
'    hndl_log.init
'    hndl_log.str_path = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\log\"
'    hndl_log.str_file_name = "log.xlsx"
'    hndl_log.open_data

    bin.init
End Function

Public Function tear_down()
'    hndl_log.close_data
End Function

Public Function test_get_place_grp_level_aisle()
    Dim specs As SpecSuite
    Dim reporter As ImmediateReporter
    
    setup
    
    Set specs = New SpecSuite
    Set reporter = New ImmediateReporter
    reporter.ListenTo specs
    
    specs.Description = "Test - get_place_grp_level_aisle"
    With specs.It("For each BIN is found correct place group - detail aisle")
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-01-13-015-10")).ToEqual "VNA_RACK"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-01-13-020-22")).ToEqual "VNA_RACK"
        
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-01-13-889-01")).ToEqual "VNA_BULK"
        
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-33-23-999-01")).ToEqual "TA_BULK"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-33-30-002-02")).ToEqual "TA_RACK"
        
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-20-02-01-007-01")).ToEqual "HBW_GATE"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-20-02-01-042-01")).ToEqual "HBW_ROBOT_IN"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-20-02-01-374-01")).ToEqual "HBW_ROBOT_IN"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-20-02-01-060-01")).ToEqual "HBW_ROBOT_OUT"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-20-02-01-376-01")).ToEqual "HBW_ROBOT_OUT"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-21-01-02-001-01")).ToEqual "HBW_WH"
        
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-20-02-01-072-01")).ToEqual "HBW_CONVEYOR_IN"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-02-20-01-013-01")).ToEqual "HBW_CONVEYOR_OUT"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-13-20-02-024-01")).ToEqual "HBW_CONVEYOR_OUT"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-20-02-020-01")).ToEqual "HBW_CONVEYOR_OUT"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-20-03-234-01")).ToEqual "HBW_CONVEYOR_OUT"
        
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-80-80-033-01")).ToEqual "RA_INBOUND"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-80-90-043-01")).ToEqual "RA_OUTBOUND"
        
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-12-03-04-921-01")).ToEqual "PROD_LINE_OUT"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-13-03-01-060-01")).ToEqual "PROD_LINE_OUT"
        .Expect(bin_place_grp.get_place_grp_level_aisle("6-13-03-01-998-01")).ToEqual "PROD_HALL"
        .Expect(bin_place_grp.get_place_grp_level_aisle("143706")).ToEqual "PROD_LINE_IN"
    End With
    
    tear_down
End Function

Public Function test_process()
'    Dim dbl_start As Double
'    Dim dbl_end As Double
'
'    setup
'
'    dbl_start = Now
'    Debug.Print "6-12-01-13-015-10 > " & (bin_place_grp.get_place_grp("6-12-01-13-015-10") = "VNA_RACK")
'    Debug.Print "6-12-01-13-020-22 > " & (bin_place_grp.get_place_grp("6-12-01-13-020-22") = "VNA_RACK")
'
'    Debug.Print "6-12-01-13-889-01 > " & (bin_place_grp.get_place_grp("6-12-01-13-889-01") = "VNA_INBOUND")
'
'    Debug.Print "6-12-33-23-999-01 > " & (bin_place_grp.get_place_grp("6-12-33-23-999-01") = "TA_INBOUND")
'    Debug.Print "6-12-33-30-002-02 > " & (bin_place_grp.get_place_grp("6-12-33-30-002-02") = "TA_RACK")
'
'    Debug.Print "6-20-02-01-007-01 > " & (bin_place_grp.get_place_grp("6-20-02-01-007-01") = "HBW_GATE")
'    Debug.Print "6-20-02-01-042-01 > " & (bin_place_grp.get_place_grp("6-20-02-01-042-01") = "HBW_ROBOT_IN")
'    Debug.Print "6-20-02-01-374-01 > " & (bin_place_grp.get_place_grp("6-20-02-01-374-01") = "HBW_ROBOT_IN")
'    Debug.Print "6-20-02-01-060-01 > " & (bin_place_grp.get_place_grp("6-20-02-01-060-01") = "HBW_ROBOT_OUT")
'    Debug.Print "6-20-02-01-376-01 > " & (bin_place_grp.get_place_grp("6-20-02-01-376-01") = "HBW_ROBOT_OUT")
'    Debug.Print "6-21-01-02-001-01 > " & (bin_place_grp.get_place_grp("6-21-01-02-001-01") = "HBW_WH")
'    Debug.Print "6-20-02-01-072-01 > " & (bin_place_grp.get_place_grp("6-20-02-01-072-01") = "HBW_CONVEYOR_IN")
'    Debug.Print "6-02-20-01-013-01 > " & (bin_place_grp.get_place_grp("6-02-20-01-013-01") = "HBW_CONVEYOR_OUT")
'    Debug.Print "6-13-20-02-024-01 > " & (bin_place_grp.get_place_grp("6-13-20-02-024-01") = "HBW_CONVEYOR_OUT")
'    Debug.Print "6-12-20-02-020-01 > " & (bin_place_grp.get_place_grp("6-12-20-02-020-01") = "HBW_CONVEYOR_OUT")
'    Debug.Print "6-12-20-03-234-01 > " & (bin_place_grp.get_place_grp("6-12-20-03-234-01") = "HBW_CONVEYOR_OUT")
'
'    Debug.Print "6-12-80-80-033-01 > " & (bin_place_grp.get_place_grp("6-12-80-80-033-01") = "RA_INBOUND")
'    Debug.Print "6-12-80-90-043-01 > " & (bin_place_grp.get_place_grp("6-12-80-90-043-01") = "RA_OUTBOUND")
'
'
'    Debug.Print "6-12-03-04-921-01 > " & (bin_place_grp.get_place_grp("6-12-03-04-921-01") = "PROD_LINE_OUT")
'    Debug.Print "6-13-03-01-060-01 > " & (bin_place_grp.get_place_grp("6-13-03-01-060-01") = "PROD_LINE_OUT")
'    Debug.Print "6-13-03-01-998-01 > " & (bin_place_grp.get_place_grp("6-13-03-01-998-01") = "PROD_HALL")
'    Debug.Print "143706 - 16B > " & (bin_place_grp.get_place_grp("143706") = "PROD_LINE_IN")
'
'
'    dbl_end = Now
'    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
'
'    tear_down
End Function

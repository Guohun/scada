Attribute VB_Name = "Module2"
Type elec_struct
 name As String
 str As String * 7
 run As Integer
 run_set As String * 1
 check As Integer
 check_set As String * 1
 run_value As String
 check_value As String
 time As Integer
 flag As String
 value As String
 cnt As String
End Type
Global elec_in(40) As elec_struct
Global china_english
Global elec_send_flag As Integer





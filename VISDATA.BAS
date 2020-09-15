Attribute VB_Name = "Module1"
Global gsDBName As Database                                     '全局库名
Global mrecFormRecordset As Recordset               '全局记录名
Global Make_Mathine    As Integer                               '全局机号
Global China_English As Boolean                                 '全局中英文标志
Public Make_Row As Integer                                          '当前所用的记录
Public Prev_run_flag    As Boolean                                 '上次运行否
Public Time_Simpling   As Single                            '每秒读入数据个数
Public Send_File_Port As Long     '送文件端口
Public Get_File_Port As Long     '收文件端口


'以下查询菜单用
Public gbFindFailed As Boolean
Public gsFindFiel     As String
Public gsFindExpr     As String
Public gsFindOp      As String
Public gnFindType As Integer
Public mbNotFound As Boolean
Public lstI As Integer




Global password As String                     '密码
Global user_name As String                  '用户名
Global Multi_Screen  As String              '多屏卡
Global Simulate_Hwnd     As Long        '模拟句柄
Global Run_Hwnd          As Long        '运行句柄

'菜单管理用
Global mnudata As Database
Global recset As Recordset
Global gsBaseRecordSet As String

Global day_time   As String             '日期
Global Grap_Name   As String        '图名
Global Data_Path   As String        '库路径名


'窗口管理用
Global Get_Focus_Fore_Color As Long
Global Los_Focus_Fore_Color As Long
Type FormState
    Deleted As Boolean
    Hwnd      As Long
    Color As Long
End Type
Public FState(40)  As FormState


'全局常量
Global Const CHINA = True
Global Const ENGLISH = False
Global Const CHANGE = True
Global Const UNCHANGE = False
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_DESKTOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&


'电信号定义
Type elect_send_type
    Flag  As Integer
    Length  As Integer
    Note_Name(0 To 500) As Integer
   data(0 To 500) As Integer
End Type
Type Fen_Disp_Dou
    Mathine As Long
    Dou     As Long
    Kind    As Long
    Send_data As Long
    Now_data  As Long
    One     As Single
    Two     As Single
    Power   As Single
    My_time As Long
    Statu   As String * 40
End Type

'通讯用
Global Comm_Port(20) As Integer
Global Plc_Port(3)    As Integer

'生产控制标志
Global Run_flag As Long
Global make_flag As Integer '生产控制标志
Public Fen_Run_flag As Long
Public Prev_Fen_Run_flag As Long

Global Elec_Input     As elect_send_type        '灯
Global Elec_Output      As elect_send_type   '阀

'动态连接库声明
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function GetCurrentThread Lib "kernel32" () As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Declare Function GetFocus Lib "user32" () As Long
Declare Function Write_VbKey Lib "z_lj_dll.dll" (ByVal X As Long) As Long
Declare Function read_elec_input Lib "z_lj_dll.dll" (ByRef X As elect_send_type) As Long
Declare Function read_elec_output Lib "z_lj_dll.dll" (ByRef Y As elect_send_type) As Long
Declare Function write_elec_input Lib "z_lj_dll.dll" (ByRef X As elect_send_type) As Long
Declare Function write_elec_output Lib "z_lj_dll.dll" (ByRef Y As elect_send_type) As Long
Declare Function Init_P Lib "z_lj_dll.dll" () As Long
Declare Function Init_V Lib "z_lj_dll.dll" () As Long
Declare Function P_DO Lib "z_lj_dll.dll" () As Long
Declare Function V_DO Lib "z_lj_dll.dll" () As Long
Declare Function Set_multi_screen Lib "z_lj_dll.dll" (ByVal Flag As Long) As Integer
Declare Function VBwrite_multi_screen Lib "z_lj_dll.dll" () As Integer
Declare Function write_Fs_command Lib "z_lj_dll.dll" (ByVal Flag As Long) As Long
Declare Function read_Fs_command Lib "z_lj_dll.dll" (ByRef Flag As Long) As Long
Declare Function read_fen_statu Lib "z_lj_dll.dll" (ByRef Flag As Fen_Disp_Dou) As Long
Declare Function write_grap Lib "z_lj_dll.dll" (ByVal Name As String) As Long
'Declare Function Init_Font Lib "z_lj_dll.dll" (ByVal X As Long, ByVal Name As String) As Long
'Declare Function write_multi_screen Lib "z_lj_dll.dll" (ByVal row As Long, ByVal col As Long, ByVal str As String) As Long


'
'错误报警输出
Static Function Speak_Error(ByVal error_num As String) As Integer
  Dim db As Database
  Dim rec As Recordset
  'Set db = OpenDatabase("c:\ljxt\ljxt.mdb")
'  Set rec = db.OpenRecordset("error_num")
 ' If rec.RecordCount = 0 Then
  '      MsgBox "错误显示程序本身错误"
   '     rec.Close
    '    db.Close
     '   Exit Function
  'End If
'      If (error_num < "0") And (error_num > td.Field.count) Then
        MsgBox error_num
        Speak_Error = 0
 '   Else
        Speak_Error = 1
  '  End If
    'tb.Close
   ' rec.Close
    'db.Close
End Function

'------------------------------------------------------------
'returns "" if a field is Null
'------------------------------------------------------------
Function vFieldVal(rvntFieldVal As Variant) As Variant
  If IsNull(rvntFieldVal) Then
    vFieldVal = gsNULL_STR
  Else
    vFieldVal = CStr(rvntFieldVal)
  End If
End Function
'------------------------------------------------------------
'returns  设定宽度的字符  多出则去掉,右对齐
'------------------------------------------------------------
Function SetRightString(rvntFieldVal As Variant, str_len As Integer) As Variant
    Dim X    As String
    'x = space(str_len)
    If str_len > Len(Trim(rvntFieldVal)) Then
            X = space(str_len - Len(Trim(rvntFieldVal)))
            X = X + Trim(rvntFieldVal)
    End If
    If str_len <= Len(Trim(rvntFieldVal)) Then
            X = Mid(Trim(rvntFieldVal), 1, str_len)
    End If
     SetRightString = X
End Function
'
'初始化电信号
Public Sub Init_Elec()
    Dim i As Integer
    Dim data   As Database
    Dim rec    As Recordset
    
    Set data = OpenDatabase("c:\ljxt\ljxt.mdb")
    Set rec = data.OpenRecordset("select * from  node_table where in_out=0")
    i = 0
    Do While Not rec.EOF()
            Elec_Input.Note_Name(i) = Val(rec.Fields("node_name").Value)
            Elec_Input.data(i) = 0
            rec.MoveNext
            i = i + 1
    Loop
    Elec_Input.Length = i - 1
    Elec_Input.Flag = 0
    rec.Close
    Call write_elec_input(Elec_Input)
    

    Set rec = data.OpenRecordset("select * from  node_table    where  in_out=1 ")
    i = 0
    Do While Not rec.EOF()
            Elec_Output.Note_Name(i) = Val(rec.Fields("node_name").Value)
            Elec_Output.data(i) = 0
            rec.MoveNext
            i = i + 1
    Loop
    Elec_Output.Length = i - 1
    Elec_Output.Flag = 0
    rec.Close
    Call write_elec_output(Elec_Output)
    data.Close
End Sub

'修改中英文
Public Sub Change_Language()
     Dim i%
     For i% = 0 To 40
        If FState(i%).Hwnd <> 0 Then
                Call PostMessage(FState(i%).Hwnd, WM_PAINT, GetDC(FState(i%).Hwnd), 1)
        End If
     Next i%
End Sub
'
'窗口管理函数
Public Function Add_Win(Hwnd As Long) As Integer
        Dim i%
        For i% = 0 To 40
                If FState(i%).Hwnd = 0 Then
                        FState(i%).Hwnd = Hwnd
                        Exit For
                End If
        Next i%
        Add_Win = 1
        MAIN_MDI!Button_Picture.Visible = False
End Function

'
'窗口管理函数
Public Function Del_Win(Hwnd As Long) As Integer
        For i% = 0 To 40
                If FState(i%).Hwnd = Hwnd Then
                        FState(i%).Hwnd = 0
                        Exit For
                End If
        Next i%
       Del_Win = 0
        For i% = 0 To 40
                If FState(i%).Hwnd <> 0 Then
                        Exit Function
                End If
        Next i%
       MAIN_MDI!Button_Picture.Visible = True
End Function

'system.ini 组成：
' 启动语音  temp_chin_english-->china_english
'当前机组   make_mathine
'图库名    grap_name
'多屏卡   multi_screen  =0 无 =1 原  =2 上下两屏
'网络路径  data_path
'鼠标     Have_Mouse
'读图库数据时间  Time_Simpling  '每秒读入数据个数
'Send_File_Port                 送文件断口
'Get_File_Port                  收文件断口
'模拟初始化 不用
Public Sub Read_System()
Dim FileNumber As Integer
Dim temp_china_english As Integer
Dim Have_Mouse As Integer
Prev_run_flag = False
On Error GoTo file_err
     FileNumber = FreeFile
    Open "c:\ljxt\system.ini" For Input As #FileNumber
     Input #FileNumber, temp_china_english
     Input #FileNumber, Make_Mathine
     Input #FileNumber, Grap_Name
     Input #FileNumber, Multi_Screen
     Input #FileNumber, Data_Path
     Input #FileNumber, Have_Mouse
     Input #FileNumber, Time_Simpling
     Input #FileNumber, Send_File_Port
     Input #FileNumber, Get_File_Port
     Close #FileNumber
     China_English = temp_china_english
     If Have_Mouse = 0 Then
        Call ShowCursor(0)
     End If
     If Multi_Screen <> 0 And Multi_Screen <> 1 And Multi_Screen <> 2 Then
            Call Speak_Error("多屏卡设置有错,默认 无")
            Multi_Screen = 0
     End If
     Call Set_multi_screen(Multi_Screen)
     Call write_grap(Grap_Name)
     Make_Row = 0
     If Make_Mathine = 0 Then
            Fen_Run_flag = 0
            Prev_Fen_Run_flag = 0
     End If
     Exit Sub
file_err:
    Call Speak_Error("系统文件system.ini有错")
    
End Sub

'
'显示录放音错误
Sub DispWaveErrorMessageBox()
    Dim msg As String
    Select Case Err
        Case conMCIErrCannotLoadDriver
            msg = "Error load media device driver."
        Case conMCIErrDeviceOpen
            msg = "The device is not open or is not known."
        Case conMCIErrInvalidDeviceID
            msg = "Invalid device id."
        Case conMCIErrInvalidFile
            msg = "Invalid filename."
        Case conMCIErrUnsupportedFunction
            msg = "Action not available for this device."
        Case Else
            msg = "Unknown error (" + str$(Err) + ")."
    End Select

    MsgBox msg, 48, conMCIAppTIitle
End Sub

'检查浮点数据
Function Check_Float(rvntVal As Variant) As Boolean
  Dim i As Integer
  Dim Dot_Data%
  Dim X   As Integer
  
  If Len(rvntVal) = 0 Or IsNull(rvntVal) Then
             Check_Float = False
             Exit Function
  End If
  Dot_Data = 0
  For i = 1 To Len(rvntVal)
     X = Asc(Mid(rvntVal, i, 1))
    If (X > Asc("9") Or X < Asc("0")) And X <> Asc(".") Then
             Check_Float = False
             Exit Function
    End If
        If X = Asc(".") Then Dot_Data = Dot_Data + 1
        If Dot_Data > 1 Then
                Check_Float = False
                Exit Function
        End If
  Next i
  Check_Float = True
End Function


'Pei_number  配方代号
'Add_time    加料时间
'Wna             秤名
'Input_Weight     准备的重量
'  返回  : >0  超重数值   =-1  wna  没设置

Public Function Check_Weight(Mathine As Integer, pei_number As String, Add_time As Integer, Wna As Integer, Input_Weight As Single, Cai_number As String, Sort_idx As Integer) As Integer
            Dim db   As Database
            Dim CommDB  As Database
            Dim rst    As Recordset     ' wna
            Dim CommRst    As Recordset     'comm
            Dim RSqlStr    As String
            Dim CommSqlStr    As String
            
                Set db = OpenDatabase(Data_Path & "\ljxt.mdb")
                Set CommDB = OpenDatabase(Data_Path & "\comm.mdb")
                Select Case Wna
                    Case 1
                        RSqlStr = "select  sum(weight) AS expr1000  from tan_hei_table  where pei_number='" _
                                    & pei_number & "'  and add_time=" & Add_time & " and mathine= " & Mathine _
                                    & "  and  ( sort_idx<>" & Sort_idx & "  ) "
'                                    & "  and  (sort_idx<>" & Sort_idx & "  or  cai_number<>'" & Cai_number & "') "
                        CommSqlStr = "select  [max],name  from send_table  where  " _
                                    & "   Code = " & Wna & " And Mathine = " & Mathine
                    Case 2
                        RSqlStr = "select  sum(weight) AS expr1000  from you_liao_table  where wna='2' and pei_number='" _
                                    & pei_number & "'  and add_time=" & Add_time & " and mathine= " & Mathine _
                                                        & "  and  ( sort_idx<>" & Sort_idx & "  ) "
                '                    & "  and  ( sort_idx<>" & Sort_idx & "  or  cai_number<>'" & Cai_number & "') "
                        CommSqlStr = "select  [max],name  from send_table  where  " _
                                    & "   Code = " & Wna & " And Mathine = " & Mathine
                    
                    Case 3
                        RSqlStr = "select  sum(weight) AS expr1000  from you_liao_table  where wna='3' and pei_number='" _
                                    & pei_number & "'  and add_time=" & Add_time & " and mathine= " & Mathine _
                                                        & "  and  ( sort_idx<>" & Sort_idx & "  ) "
'                                    & "  and  (sort_idx<>" & Sort_idx & "  or  cai_number<>'" & Cai_number & "') "
                        CommSqlStr = "select  [max],name  from send_table  where  " _
                                    & "   Code = " & Wna & " And Mathine = " & Mathine
                    
                    Case 4
                        RSqlStr = "select  sum(weight) AS expr1000  from jiao_liao_table  where pei_number='" _
                                    & pei_number & "'  and add_time=" & Add_time & " and mathine= " & Mathine _
                                                        & "  and  ( sort_idx<>" & Sort_idx & "  ) "
'                                    & "  and  (sort_idx<>" & Sort_idx & "  or  cai_number<>'" & Cai_number & "') "
                        CommSqlStr = "select  [max],name  from send_table  where  " _
                                    & "   Code = " & Wna & " And Mathine = " & Mathine
                    
                End Select
                
                Debug.Print RSqlStr
                Set rst = db.OpenRecordset(RSqlStr)
                Set CommRst = CommDB.OpenRecordset(CommSqlStr)
                If CommRst.EOF() Then
                                Call Speak_Error(Trim(CommRst.Fields("name").Value) + " 无设置")
                                Check_Weight = -1
                                GoTo Exit_Fun
                End If
                Debug.Print CommRst.Fields("max").Value
                Debug.Print rst.Fields("expr1000").Value
                If Not rst.EOF() Then
                           If IsNull(rst.Fields("expr1000").Value) Then
                            If (Input_Weight > CommRst.Fields("max").Value) Then
                                Call Speak_Error(Trim(CommRst.Fields("name").Value) & " 超重" & (Input_Weight - CommRst.Fields("max").Value))
                                Check_Weight = Input_Weight - CommRst.Fields("max").Value
                                GoTo Exit_Fun
                            End If
                           Else
                                If (Val(rst.Fields("expr1000").Value)) + Input_Weight > CommRst.Fields("max").Value Then
                                    Call Speak_Error(Trim(CommRst.Fields("name").Value) & " 超重" & (rst.Fields("expr1000").Value + Input_Weight - CommRst.Fields("max").Value))
                                    Check_Weight = rst.Fields("expr1000").Value + Input_Weight - CommRst.Fields("max").Value
                                    GoTo Exit_Fun
                                End If
                           End If
                End If
                Check_Weight = 0
Exit_Fun:
                rst.Close
                CommRst.Close
                db.Close
                CommDB.Close
End Function


'
'删除生产统计数据
'删除温度，能量，及生产数据

 Public Function Del_Make_Data(Date_Value As Date)
    Dim dbsNorthwind As Database
    Dim strSQLChange As String
    Dim strsqlrestore As String
    Dim qdfChange As Recordset
    Dim rstEmployees As Recordset
    Dim errLoop As Error
     Dim Temp_Value As Long
     
    ' Define two SQL statements for action queries.
    strSQLChange = " select only_key  from make_index as  a ,make_table AS b " _
                           & " WHERE a.sort_idx=b.sort_idx and a.mathine=b.mathine   and b.m_date<=#" & Date_Value & "#"


    Set dbsNorthwind = OpenDatabase(Data_Path & "\make.mdb")
    Set qdfChange = dbsNorthwind.OpenRecordset(strSQLChange, dbOpenForwardOnly)
    Do While Not qdfChange.EOF()
            Temp_Value = qdfChange.Fields(0).Value
            strsqlrestore = "delete  from make_lian_liao where  " & _
                " only_key=" & Temp_Value
            dbsNorthwind.Execute strsqlrestore, dbFailOnError
            strsqlrestore = "delete  from make_pei_fan where  " & _
                " only_key=" & Temp_Value
           dbsNorthwind.Execute strsqlrestore, dbFailOnError
            strsqlrestore = "delete  from make_base where  " & _
                " only_key=" & Temp_Value
            dbsNorthwind.Execute strsqlrestore, dbFailOnError
            qdfChange.MoveNext
    Loop
    On Error GoTo Err_Execute
    
    qdfChange.Close
    
        strsqlrestore = "delete  from make_table where  " & _
                " m_date<=#" & Date_Value & "#"
        dbsNorthwind.Execute strsqlrestore, dbFailOnError
        strsqlrestore = "delete  from make_index where  " & _
                " begin_date<=#" & Date_Value & "#"
        dbsNorthwind.Execute strsqlrestore, dbFailOnError
        
        dbsNorthwind.Close
    Exit Function
Err_Execute:
    Speak_Error ("删除错误")
End Function



Public Sub Check_KeyCode(KeyCode As Integer, Shift As Integer)
   If Shift = 4 Then
            Select Case KeyCode
                    Case vbKeyF
                                Call Write_VbKey(100)
                    Case vbKeyE
                               Call Write_VbKey(101)
                    Case vbKeyX
                                Call Write_VbKey(102)
                    Case vbKeyT
                                Call Write_VbKey(103)
                    Case vbKeyY
                                Call Write_VbKey(104)
                    Case vbKeyJ
                                Call Write_VbKey(105)
                    Case vbKeyK
                                Call Write_VbKey(106)
                    Case vbKeyM
                                Call Write_VbKey(107)
            End Select
    End If


End Sub

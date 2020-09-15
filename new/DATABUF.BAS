Attribute VB_Name = "databuf"
'                  图形缓冲区规范书
'一 , 说明:
'      用于控制程序与通讯子程序；控制程序与运行程序，报警程序之间的关系 。
'二 , 规范:
'全局变量:
'Global Const CHINA = True
'Global Const ENGLISH = False
'Global china_english As Boolean
'Global Run_flag As Integer

Type elec_in_table
            name                   As String * 10                 '电信号名
            str                        As String * 7                 '保留--->动作
            run        As Integer                   '起动接点
            run_set As Integer                  '逻辑设定值
                 
            check As Integer                '检测接点
            check_set As Integer        '逻辑设定值
            run_value As Integer
            check_value As Integer
            check_time        As Integer               '检测时间
            flag         As Integer     '停检标志 (保留)
            Value      As Integer
            count      As Integer            '计数 (保留)
            error_name   As String * 10    '错误代号
End Type
 '               如 {"上料栓","降",1001,1,1002,0,0,1,10,0,0,"111＂}


'      电信号缓冲区(read only)            每0.5S一次
'        Type elect_send_type
 '           elec_in     As Integer                      '=1  正在写   =0  可读
  '          Length      As Integer                      '实际长度
   '         Note_Name(100)   As Integer   '节点名
    '        data(100)   As Integer          '数据
  '      End Type
  
'        Global elect_receive_input       As elect_send_type         '启动节点
'        Global elect_receive_output       As elect_send_type        '检测节点

'         数据缓冲区(timing  read only)

'             If duan_flag <> 2 Then
'                    Exit Sub
'             End If
'            每生产一段
            Type make_lian_liao
                  duan_write      As Integer            ' 写段标志  ＝0 可读   ＝1 在写    用于 中间程序
                  duan_hao       As Integer
                  set_hun_time   As Integer          '/* 设定混料时间 */
                  get_hun_time   As Integer          '/* 实际混料时间 */
                  set_you_time   As Integer             '/* 所投油    为0 不投*/
                  set_tan_time   As Integer        '/* 所投炭   为 0 不投*/
                  set_jiao_time  As Integer        '/* 所投胶  为  0 不投*/
                  set_xiao_time  As Integer       ' /* 所投小料  为 0 不投*/
                  get_you_time   As Integer        '/* 所投油    为0 不投*/
                  get_tan_time   As Integer       '/* 所投炭   为 0 不投*/
                  get_jiao_time  As Integer       '/* 所投胶  为  0 不投*/
                  get_xiao_time As Integer      ' /* 所投小料  为 0 不投*/

                  next_tempro     As Integer
                  next_power      As Integer

            End Type
            Global make_input  As make_lian_liao

            '每生产完一车
            'If che_flag <> 2 Then
             '   exit
            'End If

            Type make_pei_fan
                   pei_number      As String * 6
                   give_time    As Integer
                   duan_hao     As Integer
                   cai_number   As String * 6
                   cai_name     As String * 15
                   mathine     As Integer          '/*机组号    0    1 2 */
                   wna         As String * 5            '/* 秤名  ban1       炭黑枰1*/
                                                                                            'ban2 炭黑枰2
                                                                                            'ban3 油料秤
                                                                                            'ban4 胶料秤
                    dou           As String * 5        '/* 斗号     dou1
                                                                            '斗1 dou2
                                                                            '斗2 dou3
                   set_weight   As Integer            '/* 标准重量 */
                   set_gon_cha  As Integer           ' /* 设定误差 */
                   get_weight   As Integer            '/* 实际重量 */
            End Type
'            eg:   1  "a1" 0,"1","rr",1,炭黑枰2,斗5,10,10,39
'            eg:   1  "a1" 0,"1","rr",1,炭黑枰2,斗6,10,10,12

Attribute VB_Name = "Module1"
Option Explicit
'接口变量
Public china_english As Boolean
'---------------------------------------------------------------------
Type gd_type
  color As Long
  width As Integer
End Type
Public tk_backcolor As Integer
Public gqg As gd_type
Public gyg As gd_type
Public glg As gd_type
 
 Public j_num As Integer         '胶料管代码
 Public q_num As Integer        '供汽管代码
 Public y_num As Integer        '供油管代码
 Public x_num As Integer          '信号灯代码
 Public o_num As Integer        '其他代码
'---------------------------------------------------------------------

Type xy_type
  left As Long
  top As Long
  move_pic_index As Integer
End Type

Type pic_type
  code As String * 6
  name As String * 20
  style As Integer
  position As Integer
  Index As Integer
'  pic_index As Image
End Type

Type xhd_type
  code As String * 6
  name As String * 20
  Index As Integer
 ' xhd_index As Integer
End Type

Type txt_type
  code As String * 6
  name As String * 20
  style As Integer          '=1  状态信息  =0  标注文字
  relation_code As String * 6
  Index As Integer
'  pic_index As Image
End Type

Type zline_type
  code As String * 6
  name As String * 20
  style As Integer
  Index As Integer
  start_pic_index As Integer
  end_pic_index As Integer
  first_line_index As Integer
  last_line_index As Integer
End Type

'use tmp_line control
Type temp_zline_type
  code As String * 6
  name As String * 20
  Index As Integer
  start_pic_index As Integer
  end_pic_index As Integer
'  zxy(0 To 100, 0 To 1) As Long
  length As Integer
  cur_zx As Long
  cur_zy As Long
End Type
Public temp_zline As temp_zline_type

Public dev_pic(0 To 100) As pic_type
Public dev_xhd(0 To 100) As xhd_type
Public dev_zline(0 To 100) As zline_type
Public dev_txt(0 To 200) As txt_type

Public pic_num As Integer              '图象数(1 - pic_num)
Public xhd_num As Integer             '信号灯数(1 - xhd_num)
Public zline_num As Integer             '折线数(0 - zline_num-1)
Public txt_num As Integer               '标注文字数

Public line_num As Integer              '所有折线的直线的总数
Public selected_pic, selected_xhd, selected_zline As Integer '被选中的图象及信号灯及折线

Public cur_dev_code As Integer
Public cur_filename As String
Public cur_command As String
Public cur_tool As Integer              '当前的设备

Public cur_zx, cur_zy As Long           '当前折点

Public cur_draw_line_direct As Boolean  '当前画线方向

Public cur_left, cur_top As Long   '当前鼠标坐标(相对于整个画板)
Public last_left, last_top As Long   '上次单击坐标(相对于整个画板)
Public offset_left, offset_top As Long '相对于设备的偏移坐标

Public cur_drag_xhd As Integer
Public cur_drag_zline As Integer
Public cur_drag_line As Integer
Public cur_drag_txt As Integer
Public cur_drag_status_relation_code As String * 6
Public cur_drag_pic As xy_type

'计算折点X,Y(顺时针:direct=true; 逆时针:direct=false)
Public Sub calc_zxy(direct As Boolean, x, y, x1, y1, x2, y2 As Integer)
 If direct = True Then
   If (x2 >= x1 And y2 > y1) Or (x2 <= x1 And y2 < y1) Then
     x = x2
     y = y1
   End If
   If (x2 < x1 And y2 >= y1) Or (x2 > x1 And y2 <= y1) Then
     x = x1
     y = y2
   End If
 Else
   If (x2 < x1 And y2 >= y1) Or (x2 > x1 And y2 <= y1) Then
     x = x2
     y = y1
   End If
   If (x2 <= x1 And y2 < y1) Or (x2 >= x1 And y2 > y1) Then
     x = x1
     y = y2
   End If
 End If
End Sub



'生成设备代码
Sub make_dev_code(dev_style As Integer, dev_code As String)
  Select Case dev_style
    Case 2
        dev_code = "X" + Format$(x_num, "00000")
        x_num = x_num + 1
    Case 1, 3, 4, 5, 6, 7, 8, 9, 10, 14
        dev_code = "O" + Format$(o_num, "00000")
        o_num = o_num + 1
    Case 11
        dev_code = "Q" + Format$(q_num, "00000")
        q_num = q_num + 1
    Case 12
        dev_code = "Y" + Format$(y_num, "00000")
        y_num = y_num + 1
    Case 13
        dev_code = "J" + Format$(j_num, "00000")
        j_num = j_num + 1
  Case Else
  End Select
  
End Sub

'计算折点X,Y(tango draw line mode)
Public Sub tango_calc_zxy(x, y, x1, y1, x2, y2 As Long)
 If x2 - x1 <> 0 Then
  If (Abs(y2 - y1) > Abs(x2 - x1)) And (((y2 - y1) / (x2 - x1)) > 1 Or ((y2 - y1) / (x2 - x1)) < -1) Then
     x = x1
     y = y2
  End If
  If (Abs(y2 - y1) <= Abs(x2 - x1)) And val_comp((y2 - y1) / (x2 - x1), -1, 1) Then
     If y2 <> y1 Then
       x = x2
       y = y1
     Else
       x = (x1 + x2) / 2
       y = y1
     End If
   End If
  Else
    x = x1
    y = (y1 + y2) / 2
  End If
End Sub




Public Function val_comp(val, x1, x2 As Long) As Boolean
Dim tmp As Long
Dim xx1, xx2
  xx1 = x1: xx2 = x2
  If x1 > x2 Then
    tmp = xx1
    xx1 = xx2
    xx2 = tmp
  End If
  If (val >= xx1) And (val <= xx2) Then
    val_comp = True
  Else
    val_comp = False
  End If
End Function




Function dev_pic_style(Index As Integer) As String
        Select Case Index
           Case 1
             dev_pic_style = "阀门"
           Case 2
             dev_pic_style = "信号灯"
           Case 3
             dev_pic_style = "人工投料斗"
           Case 4
             dev_pic_style = "日储斗"
           Case 5
             dev_pic_style = "油料秤"
           Case 6
             dev_pic_style = "胶料秤"
           Case 7
             dev_pic_style = "碳黑秤"
           Case 8
             dev_pic_style = "碳黑库"
           Case 9
             dev_pic_style = "炼胶炉"
           Case 10
             dev_pic_style = "供胶机"
           Case 11
             dev_pic_style = "供汽管"
           Case 12
             dev_pic_style = "供油管"
           Case 13
             dev_pic_style = "供料管"
           Case 14
             dev_pic_style = "皮带"
           Case 15
             dev_pic_style = "标注文字"
           Case 16
             dev_pic_style = "状态信息"
           Case Else
        End Select
End Function




Attribute VB_Name = "RunTkBas"
'管道流速常量
Public Const LS_STOP = 0
Public Const LS_SLOW = 1
Public Const LS_NORMAL = 2
Public Const LS_FAST1 = 3
Public Const LS_FAST2 = 4
Public Const LS_FAST3 = 5
Public Const LS_FAST4 = 6

'-------------------------------------------------------------------------------------------------
Public CT_DFS As New CTestDFS                  '测试流体算法类

Public TK_Mathine As Integer                        '当前演示图库所属机号
Public Runtk_DataBase_Path As String        '图库所在目录
Public Cur_FileName As String                       '当前演示图库名称
Public Old_Window_Width, Old_Window_Height As Long      '原窗口宽高
Public Old_Window_ScaleWidth, Old_Window_ScaleHeight As Long '原窗口宽高比例
Public Window_Scale_X As Double                 '窗口缩放比例X
Public Window_Scale_Y As Double                 '窗口缩放比例Y
Public SF_With_Status_Txt As Integer            '缩放时改变状态文字

'------------------------------------------
Type LJL_Type                                                   '密炼机类型
    Name As String * 20                                      '密炼机名称
    index As Integer                                            '密炼机索引值
    SDS_High_IDX As Integer                         '高位上顶栓索引值
    SDS_High_Style As Integer                       '高位上顶栓类型（0--信号灯，1--阀门）
    SDS_Mid_IDX As Integer                              '中位上顶栓索引值
    SDS_Mid_Style As Integer                            '中位上顶栓类型（0--信号灯，1--阀门）
    SDS_Low_IDX  As Integer                             '低位上顶栓索引值
    SDS_Low_Style As Integer                           '低位上顶栓类型（0--信号灯，1--阀门）
    
    JLM_OPEN_FM_IDX As Integer
    JLM_OPEN_FM_Style As Integer
    JLM_CLOSE_FM_IDX As Integer
    JLM_CLOSE_FM_Style As Integer
    
    XDS_OPEN_FM_IDX As Integer
    XDS_OPEN_FM_Style As Integer
    XDS_CLOSE_FM_IDX As Integer
    XDS_CLOSE_FM_Style As Integer
    
    SDS_Value As Integer
    JLM_Value As Integer
    XDS_Value As Integer
    Last_SDS_Status As Integer
    Last_JLM_Status As Integer
    Last_XDS_Status As Integer
End Type
Public LJL_Num As Integer
Public LJL(0 To 10) As LJL_Type
'------------------------------------------

Type Dash_Line_Style
    Space_len As Integer
    Xd_len As Integer
End Type
Public Dls As Dash_Line_Style
Public Dash_line_offset As Integer

Type Gd_Type
  Color As Long
  Width As Integer
  Neijing_width As Integer
  Obj_color As Long
End Type
Public Tk_backcolor As Integer
Public Gqg As Gd_Type
Public Gyg As Gd_Type
Public Glg As Gd_Type
 
 Public J_num As Integer         '胶料管代码
 Public Q_num As Integer        '供汽管代码
 Public Y_num As Integer        '供油管代码
 Public X_num As Integer          '信号灯代码
 Public O_num As Integer        '其他代码
'---------------------------------------------------------------------
'作动画用
Type Dh_FaMen_Type
    scale_x1 As Double
    scale_y1 As Double
    scale_w1 As Double
    scale_h1 As Double
    scale_x2 As Double
    scale_y2 As Double
    scale_w2 As Double
    scale_h2 As Double
End Type
Public Fm_hole(0 To 3) As Dh_FaMen_Type

'---------------------------------------------------------------------
Type Xy_Type
  Left As Long
  Top As Long
End Type

Type Pic_Type
  Code As String * 6
  Name As String * 20
  Machine As Integer
  Style As Integer
  Position As Integer
  Left As Long
  Top As Long
  Width As Long
  Height As Long
  Old_Left As Long
  Old_Top As Long
  Old_Width As Long
  Old_Height As Long
  Value As Single
  Last_Beizhu As String * 4
  Beizhu As String * 4
  Min As Single
  Max As Single
  Max_Size As Single
  Last_Status As Single
  Status As Single
  Enabled As Boolean
  Can_Comein As Boolean                            '能流进
  Can_Goout As Boolean                               '能流出
  Obj_Passing As Boolean                            '液体正在流过
   Action As Integer
End Type

Type Xhd_Type
  Code As String * 6
  Name As String * 20
  Machine As Integer
  Left As Long
  Top As Long
  Width As Long
  Height As Long
  Old_Left As Long
  Old_Top As Long
  Old_Width  As Long
  Old_Height As Long
  Last_Status As Integer
  Status As Integer
  Enabled As Boolean
  Value As Integer
   Action As Integer
End Type

Type Status_Txt_Type
  Code As String * 6
  Name As String * 20
  Left As Long
  Top As Long
  Old_Left As Long
  Old_Top As Long
  Font_size As Long
  Old_font_size As Long
  Relation_code As String * 6
  Enabled As Boolean
End Type

Type Zline_Type
  Code As String * 6
  Name As String * 20
  Style As Integer
  Start_Pic_Index As Integer
  End_Pic_Index As Integer
  Last_Status As Integer
  Status As Integer
  Visible As Boolean
  Zxy(0 To 100, 0 To 1) As Integer
  Zdot_num As Integer
  Enabled As Boolean
  Can_Comein As Boolean                             '能流进
  Can_Goout As Boolean                              '能流出
  Obj_Passing As Boolean                            '液体正在流过
  
  Cur_Offset As Single
  NJ As Integer                                 '内径
  Speed As Single                            '流速
End Type

Public Dev_PIC(0 To 100) As Pic_Type
Public Dev_XHD(0 To 100) As Xhd_Type
Public Dev_ZLine(0 To 100) As Zline_Type
Public Status_Txt(0 To 200) As Status_Txt_Type

Public Pic_num As Integer              '图象数(1 - pic_num)
Public Xhd_num As Integer             '信号灯数(1 - xhd_num)
Public Zline_num As Integer             '折线数(0 - zline_num-1)
Public Status_txt_num As Integer               '状态文字数
Public Txt_num As Integer                           '标注文字数
Public Kj_num As Integer                           '块胶状态数

Public Line_num As Integer              '所有折线的直线的总数

Public Cur_left, Cur_top As Long   '当前鼠标坐标(相对于整个画板)
Public Offset_left, Offset_top As Long '相对于设备的偏移坐标

'状态信息结构
Type Run_Table      '* vb-->dll --->VB */
        total_che   As Integer    '; /*总车数*/
        Now_Che     As Integer  ';/*第几车*/
        All_Duan        As Integer ';/* 总段 */
        ban                 As Integer  '    /*班  */
        Mathine         As Integer  ';  /*机号 */
        All_Time             As Long     ';   /*总工作时间 */
        All_Che_Time      As Long  '; /*总一车时间 */
        Duan_Time       As Long   ' /*段时间 */
        number      As String * 8 '/*配方编号 */
        Name            As String * 18 '   /*配方名称 */
End Type
Public Run_Do_Table As Run_Table
'--------------------------------------------------------------------------
'图的深度优先算法数据结构

'环路节点与派生节点的编号对应关系
Type Rollback_ZLine_Type
    ZIndex(0 To 20) As Integer
    Length As Integer
End Type
Public Rollback_ZLine As Rollback_ZLine_Type
Type Der_Vtx_Type
    old_index(0 To 20) As Integer
    Der_Index(0 To 20) As Integer
    Length As Integer
End Type
Public Der_Vtx As Der_Vtx_Type

Public Arcs(1 To 100, 1 To 100) As Integer     '图的关联矩阵，元素为折线索引值
Public Visited(1 To 100) As Boolean       '记录访问标志
Public Passed_vtx(1 To 100) As Boolean
Type road_type
    ZIndex(0 To 100) As Integer            '记录折线路径
    Length As Integer
End Type
Public Road As road_type
Public Zline_obj_passing(0 To 100) As Boolean
Public Search_new_road As Boolean
Public Last_Vtx As Integer                                 '上次访问的顶点
'============DLL=================
'Declare Function read_select_dou Lib "c:\ljxt\z_lj_dll.dll" (ByVal name As String, ByVal mathine As Integer, flag As Integer) As Integer
'Declare Function read_ban Lib "c:\ljxt\z_lj_dll.dll" (ByVal name As String, ByVal mathine As Integer, set_weight As Integer, get_weight As Integer, flag As Integer) As Integer

'============API Function==============
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
'stretch mode
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
Public Const FLOODFILLSURFACE = 1
Public Const FLOODFILLBORDER = 0
Type POINTAPI
        x As Long
        Y As Long
End Type
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long   'WINDING
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_PURGE = &H40               '  purge non-static events for task
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function write_run Lib "z_lj_dll.dll" (ByRef Y As Run_Table, ByVal Flag As Long) As Long
Declare Function read_run Lib "z_lj_dll.dll" (ByRef Y As Run_Table, ByRef Flag As Long) As Long
Declare Function read_ban Lib "z_lj_dll.dll" (ByVal x As Integer, ByVal Y As Integer, Z As Single, m As Single, temp As Integer) As Integer
Declare Function read_Tan_dou Lib "z_lj_dll.dll" (ByVal x As Integer, ByVal Y As Integer, Z As Integer) As Integer
Declare Function read_ri_chu_dou Lib "z_lj_dll.dll" (ByVal x As Integer, ByVal Y As Integer, Z As Integer) As Integer
Declare Function read_You_dou Lib "z_lj_dll.dll" (ByVal x As Integer, ByVal Y As Integer, Z As Integer) As Integer


'错误报警
Public No_Speaking  As Boolean
Public Error_Count As Integer
Public Cur_Speaking_Error_Code As Integer
Public Last_Speaking_Error_Code As Long
Public Speaking_Error, Reading_ErrorHelp As Boolean
Type G_error_num
    Length As Integer
    Code(0 To 100)  As Long
    Prompt(0 To 100)   As String * 99
End Type
Public Last_Errors As G_error_num
Public Cur_Errors As G_error_num
Declare Function get_error Lib "z_lj_dll.dll" (ByRef x As G_error_num) As Integer
Declare Function kill_error Lib "z_lj_dll.dll" (ByVal x As Long) As Integer
Declare Function read_error Lib "z_lj_dll.dll" (ByVal x As Long) As Integer
Declare Function write_error Lib "z_lj_dll.dll" (ByVal x As Long, ByVal Y As String) As Integer
Declare Function write_multi_screen Lib "z_lj_dll.dll" (ByVal row As Long, ByVal col As Long, ByVal str As String) As Integer



'Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Function PixelX(x As Long) As Integer
PixelX = CInt(x / Screen.TwipsPerPixelX)
End Function
Public Function PixelY(Y As Long) As Integer
PixelY = CInt(Y / Screen.TwipsPerPixelY)
End Function


Public Function Val_Comp(Val, X1, X2 As Long) As Boolean
Dim Tmp As Long
Dim xx1, xx2
  xx1 = X1: xx2 = X2
  If X1 > X2 Then
    Tmp = xx1
    xx1 = xx2
    xx2 = Tmp
  End If
  If (Val >= xx1) And (Val <= xx2) Then
    Val_Comp = True
  Else
    Val_Comp = False
  End If
End Function

Public Function Draw_Line(Frm As Form, X1, Y1, X2, Y2 As Integer, Start_XD As Single, Space_Color As Long, Line_color As Long)
Dim T As Double
Dim x, Y, DX, DY As Integer

If Abs(X2 - X1) > Abs(Y2 - Y1) Then     '-45du -- 45du or 135du --225du
    DY = 0: DX = IIf(X2 < X1, -1, 1):    T = (Y2 - Y1) / (X2 - X1)
Else
    If X2 = X1 Then                                     '90du or -90du
    Else                                                        '
        T = (Y2 - Y1) / (X2 - X1)
    End If
    DX = 0: DY = IIf(Y2 < Y1, -1, 1)
End If

If DX <> 0 Then
    '画起始线段
    x = X1: Y = Y1: Frm.CurrentX = x: Frm.CurrentY = Y
    If Start_XD > Dls.Xd_len Then
        x = X1 + DX * (Start_XD - Dls.Xd_len): Y = T * (x - X1) + Y1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Space_Color
        Else
            Frm.Line -(X2, Y2), Space_Color
            Exit Function
        End If
        x = x + DX * Dls.Xd_len: Y = T * (x - X1) + Y1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Line_color
        Else
            Frm.Line -(X2, Y2), Line_color
            Exit Function
        End If
    Else
        x = X1 + DX * Start_XD: Y = T * (x - X1) + Y1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Line_color
        Else
            Frm.Line -(X2, Y2), Line_color
            Exit Function
        End If
    End If
    
    
    '画空格及线段
    Do
        x = x + DX * Dls.Space_len:  Y = T * (x - X1) + Y1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Space_Color
        Else
            Frm.Line -(X2, Y2), Space_Color
            Exit Function
        End If
        
        x = x + DX * Dls.Xd_len:  Y = T * (x - X1) + Y1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Line_color
        Else
            Frm.Line -(X2, Y2), Line_color
            Exit Function
        End If
    Loop
Else    'dx=0
    If X2 = X1 Then
        '画起始线段
        x = X1: Y = Y1: Frm.CurrentX = x: Frm.CurrentY = Y
    If Start_XD > Dls.Xd_len Then
        Y = Y1 + DY * (Start_XD - Dls.Xd_len): x = X1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Space_Color
        Else
            Frm.Line -(X2, Y2), Space_Color
            Exit Function
        End If
        Y = Y + DY * Dls.Xd_len: x = X1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Line_color
        Else
            Frm.Line -(X2, Y2), Line_color
            Exit Function
        End If
    Else
        Y = Y1 + DY * Start_XD:  x = X1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Line_color
        Else
            Frm.Line -(X2, Y2), Line_color
            Exit Function
        End If
    End If
    
        '画空格及线段
        Do
            Y = Y + DY * Dls.Space_len:  x = X1
            If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
                Frm.Line -(x, Y), Space_Color
            Else
                Frm.Line -(X2, Y2), Space_Color
                Exit Function
            End If
        
            Y = Y + DY * Dls.Xd_len:  x = X1
            If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
                Frm.Line -(x, Y), Line_color
            Else
                Frm.Line -(X2, Y2), Line_color
                Exit Function
            End If
        Loop
    Else    'x2<>x1
        '画起始线段
        x = X1: Y = Y1: Frm.CurrentX = x: Frm.CurrentY = Y
    If Start_XD > Dls.Xd_len Then
        Y = Y1 + DY * (Start_XD - Dls.Xd_len): x = (Y - Y1) / T + X1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Space_Color
        Else
            Frm.Line -(X2, Y2), Space_Color
            Exit Function
        End If
        Y = Y + DY * Dls.Xd_len: x = (Y - Y1) / T + X1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Line_color
        Else
            Frm.Line -(X2, Y2), Line_color
            Exit Function
        End If
    Else
        Y = Y1 + DY * Start_XD:  x = (Y - Y1) / T + X1
        If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
            Frm.Line -(x, Y), Line_color
        Else
            Frm.Line -(X2, Y2), Line_color
            Exit Function
        End If
    End If
    
        '画空格及线段
        Do
            Y = Y + DY * Dls.Space_len:  x = (Y - Y1) / T + X1
            If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
                Frm.Line -(x, Y), Space_Color
            Else
                Frm.Line -(X2, Y2), Space_Color
                Exit Function
            End If
        
            Y = Y + DY * Dls.Xd_len:  x = (Y - Y1) / T + X1
            If Val_Comp(x, X1, CLng(X2)) And Val_Comp(Y, Y1, CLng(Y2)) Then
                Frm.Line -(x, Y), Line_color
            Else
                Frm.Line -(X2, Y2), Line_color
                Exit Function
            End If
        Loop
    End If  'x2=x1
End If  'dx<>0

End Function




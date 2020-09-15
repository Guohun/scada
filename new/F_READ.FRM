VERSION 4.00
Begin VB.Form f_read 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打开图库"
   ClientHeight    =   4356
   ClientLeft      =   2220
   ClientTop       =   1476
   ClientWidth     =   5772
   Height          =   4740
   Left            =   2172
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4356
   ScaleWidth      =   5772
   ShowInTaskbar   =   0   'False
   Top             =   1140
   Width           =   5868
   Begin Threed.SSFrame SSFrame1 
      Height          =   3972
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   3132
      _Version        =   65536
      _ExtentX        =   5525
      _ExtentY        =   7006
      _StockProps     =   14
      Caption         =   "请选择"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox Text2 
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1452
         Left            =   240
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2280
         Width           =   2532
      End
      Begin VB.TextBox Text1 
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1440
         Width           =   2532
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2532
      End
      Begin VB.Label Label3 
         Caption         =   "说    明："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "作    者："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "图库名："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   972
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   492
      Left            =   4080
      TabIndex        =   4
      Top             =   1200
      Width           =   1212
      _Version        =   65536
      _ExtentX        =   2138
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "退  出"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
   Begin Threed.SSCommand SSCommand1 
      Default         =   -1  'True
      Height          =   492
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   1212
      _Version        =   65536
      _ExtentX        =   2138
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "确  定"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
End
Attribute VB_Name = "f_read"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Dim db As Database, rs As Recordset
Public Function draw_pic(Index As Integer)
Dim Temp, H, W As Integer
    f_read.load_to_temp_pic Dev_pic(Index).Style, Dev_pic(Index).Position
    'temp = BitBlt(f_showlc.hdc, dev_pic(Index).left, dev_pic(Index).top, dev_pic(Index).width, dev_pic(Index).height, f_showlc!temp_pic.hdc, 0, 0, SRCAND)
   W = Dev_pic(Index).Width: H = Dev_pic(Index).Height
   SetStretchBltMode F_ShowLc.hdc, COLORONCOLOR
   StretchBlt F_ShowLc.hdc, Dev_pic(Index).Left, Dev_pic(Index).Top, W, H, F_ShowLc!temp_pic.hdc, 0, 0, F_ShowLc!temp_pic.Width * Window_Scale_X - 1, F_ShowLc!temp_pic.Height * Window_Scale_Y - 1, SRCAND
End Function

Public Sub load_to_temp_pic(which_bmp As Integer, pos As Integer)
        Select Case which_bmp
           Case 1
                Select Case pos
                    Case 0
                        F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\fmleft.bmp")
                    Case 1
                        F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\fmup.bmp")
                    Case 2
                        F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\fmright.bmp")
                    Case 3
                        F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\fmdown.bmp")
                End Select
             'f_showlc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\famen.bmp")
'           Case 2
'             f_showlc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\xhd.bmp")
           Case 3
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\tld.bmp")
           Case 4
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\rcd.bmp")
           Case 5       '隐含的阀门
             F_ShowLc!temp_pic.Picture = LoadPicture("")
'             f_showlc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\fmhide.bmp")
           Case 6
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\jlc.bmp")
           Case 7
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\thc.bmp")
           Case 8
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\thk.bmp")
           Case 9
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\ljl.bmp")
           Case 10
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\gjj.bmp")
           Case 14
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\pidai.bmp")
           Case 17
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\ysg.bmp")
           Case 18
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\eccd.bmp")
           Case 19
             F_ShowLc!temp_pic.Picture = LoadPicture("c:\ljxt\tkbmp\yld.bmp")
           Case Else
        End Select
End Sub

'从数据库取出设备数据
Sub load_from_mdb(Mdb_name As String)
Dim Loc_DB, Comm_Db As Database, Loc_RS, Sl_RS, Comm_RS As Recordset
Dim i, j, Zxy_index As Integer
Dim Head As String
If Trim(Mdb_name) = "" Or Dir$(Trim(Mdb_name)) = Null Then
  Exit Sub
End If
Set Loc_DB = Workspaces(0).OpenDatabase(Trim(Mdb_name))
Set Loc_RS = Loc_DB.OpenRecordset("grap_table", dbOpenDynaset)
Set rl_rs = Loc_DB.OpenRecordset("relation_table", dbOpenDynaset)
Set Sl_RS = Loc_DB.OpenRecordset("sub_line_table", dbOpenDynaset)

Kj_num = 0
Tl_num = 0
'读取图象
Pic_num = 0
Selected_pic = -1
Loc_RS.FindFirst "trim(code)>='O00000' and trim(code)<='O99999'"
Do Until Loc_RS.NoMatch
    Pic_num = Pic_num + 1
    Dev_pic(Pic_num).Code = Trim(Loc_RS!Code)
    Dev_pic(Pic_num).Name = Trim(Loc_RS!Name)
    Dev_pic(Pic_num).Machine = Loc_RS!Machine
    Dev_pic(Pic_num).Style = Loc_RS!Style
    Dev_pic(Pic_num).Position = Loc_RS!Position
    Dev_pic(Pic_num).Left = PixelX(Loc_RS!Left)
    Dev_pic(Pic_num).Top = PixelY(Loc_RS!Top)
    Dev_pic(Pic_num).Width = PixelX(Loc_RS!Width)
    Dev_pic(Pic_num).Height = PixelY(Loc_RS!Height)
    Dev_pic(Pic_num).Enabled = True
    draw_pic (Pic_num)
    If (Dev_pic(Pic_num).Style = 6) Then  '胶料秤
        Kj_num = Kj_num + 1
        Load F_ShowLc.Kj(Kj_num)
        F_ShowLc.Kj(Kj_num).Left = Dev_pic(Pic_num).Left
        F_ShowLc.Kj(Kj_num).Width = Dev_pic(Pic_num).Width
        F_ShowLc.Kj(Kj_num).Top = Dev_pic(Pic_num).Top - F_ShowLc.Kj(Kj_num).Height
        F_ShowLc.Kj(Kj_num).Tag = str(Pic_num)
        F_ShowLc.Kj(Kj_num).Visible = True
    End If
'    load_to_temp_pic (dev_pic(pic_num).style)
'    tl_num = BitBlt(f_showlc.hDC, pixelx(loc_rs!left), pixely(loc_rs!top), pixelx(loc_rs!width), pixely(loc_rs!height), f_showlc!temp_pic.hDC, 0, 0, SRCAND)
'    Load f_showlc!p(pic_num)
'    Call load_pic(pic_num, dev_pic(pic_num).style)
'    f_showlc!p(pic_num).left = pixelx(loc_rs!left)
'    f_showlc!p(pic_num).top = pixely(loc_rs!top)
'    f_showlc!p(pic_num).width = pixelx(loc_rs!width)
'    f_showlc!p(pic_num).Height = pixely(loc_rs!Height)
    
'    f_showlc!p(pic_num).visible = True
'    f_showlc!p(pic_num).ZOrder 0
    
    O_num = O_num + 1
    Loc_RS.FindNext "trim(code)>='O00000' and trim(code)<='O99999'"
Loop

'===============================================================
'读取折线(relation_table,sub_line_table)
Zline_num = 0
Selected_zline = -1
Line_num = 0
J_num = 0: Y_num = 0: Q_num = 0
'供胶管（料）
    '由relation_table查找管道
    Loc_RS.FindFirst "trim(code)>='J00000' and trim(code)<='J99999'"
    Do Until Loc_RS.NoMatch
        Dev_zline(Zline_num).Code = Trim(Loc_RS!Code)
        Dev_zline(Zline_num).Name = Trim(Loc_RS!Name)
        Dev_zline(Zline_num).Style = Loc_RS!Style
        Dev_zline(Zline_num).Enabled = True
        
        'start_pic_index,   end_pic_index
        rl_rs.FindFirst "code='" + Trim(Dev_zline(Zline_num).Code) + "'"
        If Not rl_rs.NoMatch Then
            For i = 1 To Pic_num
                If Trim(Dev_pic(i).Code) = Trim(rl_rs!last_code) Then
                    Dev_zline(Zline_num).Start_pic_index = i
                    Exit For
                End If
            Next i
            For i = 1 To Pic_num
                If Trim(Dev_pic(i).Code) = Trim(rl_rs!next_code) Then
                    Dev_zline(Zline_num).End_pic_index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        Zxy_index = 0
        Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=0"
        If Not Sl_RS.NoMatch Then
            F_ShowLc.CurrentX = PixelX(Sl_RS!Left)
            F_ShowLc.CurrentY = PixelX(Sl_RS!Top)
            Dev_zline(Zline_num).Zxy(Zxy_index, 0) = PixelX(Sl_RS!Left)
            Dev_zline(Zline_num).Zxy(Zxy_index, 1) = PixelY(Sl_RS!Top)
        End If
        
        Zxy_index = Zxy_index + 1
        Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=" + str(Zxy_index)
        Do Until Sl_RS.NoMatch
            F_ShowLc.DrawWidth = Glg.Width
            F_ShowLc.Line -(PixelX(Sl_RS!Left), PixelY(Sl_RS!Top)), Glg.Color
    
            Dev_zline(Zline_num).Zxy(Zxy_index, 0) = PixelX(Sl_RS!Left)
            Dev_zline(Zline_num).Zxy(Zxy_index, 1) = PixelY(Sl_RS!Top)
            
            Zxy_index = Zxy_index + 1
            Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=" + str(Zxy_index)
        Loop
        Dev_zline(Zline_num).Zdot_num = Zxy_index
        
        J_num = J_num + 1
        Zline_num = Zline_num + 1
        
        Loc_RS.FindNext "trim(code)>='J00000' and trim(code)<='J99999'"
    Loop
  
'供汽管
    '由relation_table查找管道
    Loc_RS.FindFirst "trim(code)>='Q00000' and trim(code)<='Q99999'"
    Do Until Loc_RS.NoMatch
        Dev_zline(Zline_num).Code = Trim(Loc_RS!Code)
        Dev_zline(Zline_num).Name = Trim(Loc_RS!Name)
        Dev_zline(Zline_num).Style = Loc_RS!Style
        Dev_zline(Zline_num).Enabled = True
        
        'start_pic_index,   end_pic_index
        rl_rs.FindFirst "code='" + Trim(Dev_zline(Zline_num).Code) + "'"
        If Not rl_rs.NoMatch Then
            For i = 1 To Pic_num
                If Trim(Dev_pic(i).Code) = Trim(rl_rs!last_code) Then
                    Dev_zline(Zline_num).Start_pic_index = i
                    Exit For
                End If
            Next i
            For i = 1 To Pic_num
                If Trim(Dev_pic(i).Code) = Trim(rl_rs!next_code) Then
                    Dev_zline(Zline_num).End_pic_index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        Zxy_index = 0
        Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=0"
        If Not Sl_RS.NoMatch Then
            F_ShowLc.CurrentX = PixelX(Sl_RS!Left)
            F_ShowLc.CurrentY = PixelY(Sl_RS!Top)
            Dev_zline(Zline_num).Zxy(Zxy_index, 0) = PixelX(Sl_RS!Left)
            Dev_zline(Zline_num).Zxy(Zxy_index, 1) = PixelY(Sl_RS!Top)
        End If
        
        Zxy_index = Zxy_index + 1
        Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=" + str(Zxy_index)
        Do Until Sl_RS.NoMatch
            F_ShowLc.DrawWidth = Glg.Width
            F_ShowLc.Line -(PixelX(Sl_RS!Left), PixelY(Sl_RS!Top)), Gqg.Color
            
            Dev_zline(Zline_num).Zxy(Zxy_index, 0) = PixelX(Sl_RS!Left)
            Dev_zline(Zline_num).Zxy(Zxy_index, 1) = PixelY(Sl_RS!Top)
            
            Zxy_index = Zxy_index + 1
            Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=" + str(Zxy_index)
        Loop
        Dev_zline(Zline_num).Zdot_num = Zxy_index
        
        Q_num = Q_num + 1
        Zline_num = Zline_num + 1
        
        Loc_RS.FindNext "trim(code)>='Q00000' and trim(code)<='Q99999'"
    Loop
  
'供油管
    '由relation_table查找管道
    Loc_RS.FindFirst "trim(code)>='Y00000' and trim(code)<='Y99999'"
    Do Until Loc_RS.NoMatch
        Dev_zline(Zline_num).Code = Trim(Loc_RS!Code)
        Dev_zline(Zline_num).Name = Trim(Loc_RS!Name)
        Dev_zline(Zline_num).Style = Loc_RS!Style
        Dev_zline(Zline_num).Enabled = True
        
        'start_pic_index,   end_pic_index
        rl_rs.FindFirst "code='" + Trim(Dev_zline(Zline_num).Code) + "'"
        If Not rl_rs.NoMatch Then
            For i = 1 To Pic_num
                If Trim(Dev_pic(i).Code) = Trim(rl_rs!last_code) Then
                    Dev_zline(Zline_num).Start_pic_index = i
                    Exit For
                End If
            Next i
            For i = 1 To Pic_num
                If Trim(Dev_pic(i).Code) = Trim(rl_rs!next_code) Then
                    Dev_zline(Zline_num).End_pic_index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        Zxy_index = 0
        Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=0"
        If Not Sl_RS.NoMatch Then
            F_ShowLc.CurrentX = PixelX(Sl_RS!Left)
            F_ShowLc.CurrentY = PixelY(Sl_RS!Top)
            Dev_zline(Zline_num).Zxy(Zxy_index, 0) = PixelX(Sl_RS!Left)
            Dev_zline(Zline_num).Zxy(Zxy_index, 1) = PixelY(Sl_RS!Top)
        End If
        
        Zxy_index = Zxy_index + 1
        Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=" + str(Zxy_index)
        Do Until Sl_RS.NoMatch
            F_ShowLc.DrawWidth = Glg.Width
            F_ShowLc.Line -(PixelX(Sl_RS!Left), PixelY(Sl_RS!Top)), Gyg.Color
            
            Dev_zline(Zline_num).Zxy(Zxy_index, 0) = PixelX(Sl_RS!Left)
            Dev_zline(Zline_num).Zxy(Zxy_index, 1) = PixelY(Sl_RS!Top)
            
            Zxy_index = Zxy_index + 1
            Sl_RS.FindFirst "trim(code)='" + Trim(Dev_zline(Zline_num).Code) + "' and index=" + str(Zxy_index)
        Loop
        Dev_zline(Zline_num).Zdot_num = Zxy_index
        
        Y_num = Y_num + 1
        Zline_num = Zline_num + 1
        
        Loc_RS.FindNext "trim(code)>='Y00000' and trim(code)<='Y99999'"
    Loop
    Sl_RS.Close
    rl_rs.Close
'===============================================================
  '读取信号灯(grap_table)
F_ShowLc.FillColor = QBColor(12)
F_ShowLc.ForeColor = QBColor(0)
F_ShowLc.DrawWidth = 1
Xhd_num = 0
Loc_RS.FindFirst "trim(code)>='X00000' and trim(code)<='X99999'"
Do Until Loc_RS.NoMatch
    Xhd_num = Xhd_num + 1
    Dev_xhd(Xhd_num).Code = Trim(Loc_RS!Code)
    Dev_xhd(Xhd_num).Name = Trim(Loc_RS!Name)
    Dev_xhd(Xhd_num).Machine = Loc_RS!Machine
    Dev_xhd(Xhd_num).Enabled = True

    Dev_xhd(Xhd_num).Left = PixelX(Loc_RS!Left)
    Dev_xhd(Xhd_num).Top = PixelY(Loc_RS!Top)
    Dev_xhd(Xhd_num).Width = PixelX(Loc_RS!Width)
    Dev_xhd(Xhd_num).Height = PixelY(Loc_RS!Height)
    Dev_xhd(Xhd_num).Old_left = PixelX(Loc_RS!Left)
    Dev_xhd(Xhd_num).Old_top = PixelY(Loc_RS!Top)
    Dev_xhd(Xhd_num).Old_width = PixelX(Loc_RS!Width)
    Dev_xhd(Xhd_num).Old_height = PixelY(Loc_RS!Height)
    Ellipse F_ShowLc.hdc, Dev_xhd(Xhd_num).Left, Dev_xhd(Xhd_num).Top, Dev_xhd(Xhd_num).Left + Dev_xhd(Xhd_num).Width, Dev_xhd(Xhd_num).Top + Dev_xhd(Xhd_num).Height
'    tmp = IIf(dev_xhd(xhd_num).width < dev_xhd(xhd_num).height, dev_xhd(xhd_num).width, dev_xhd(xhd_num).height)
'    Circle (dev_xhd(xhd_num).left + tmp \ 2, dev_xhd(xhd_num).top + tmp \ 2), tmp \ 2

'    Load f_showlc!xhd(xhd_num)
    'need code
'    f_showlc!xhd(xhd_num).left = pixelx(loc_rs!left)
'    f_showlc!xhd(xhd_num).top = pixely(loc_rs!top)
'    f_showlc!xhd(xhd_num).width = pixelx(loc_rs!width)
'    f_showlc!xhd(xhd_num).height = pixely(loc_rs!height)
    
'    f_showlc!xhd(xhd_num).visible = True
'    f_showlc!xhd(xhd_num).ZOrder 0
    
    X_num = X_num + 1
    Loc_RS.FindNext "trim(code)>='X00000' and trim(code)<='X99999'"
Loop
Loc_RS.Close
    
'读取文字 (prompt_table)
    i = 0
    Txt_num = 0
    Status_txt_num = 0
    Set Loc_RS = Loc_DB.OpenRecordset("prompt_table", dbOpenDynaset)
    Do Until Loc_RS.EOF
      If Loc_RS!Style = 0 Then
        Txt_num = Txt_num + 1
        Load F_ShowLc!txt(Txt_num)
        F_ShowLc!txt(Txt_num).Caption = Trim(Loc_RS!prompt)
        F_ShowLc!txt(Txt_num).ForeColor = Loc_RS!font_color
        On Error GoTo skip1
        F_ShowLc!txt(Txt_num).FontName = Trim(Loc_RS!font_name)
skip1:
        On Error GoTo 0
        F_ShowLc!txt(Txt_num).FontSize = Loc_RS!Font_size
        F_ShowLc!txt(Txt_num).Tag = str(F_ShowLc!txt(Txt_num).FontSize)
        F_ShowLc!txt(Txt_num).Left = PixelX(Loc_RS!Left)
        F_ShowLc!txt(Txt_num).Top = PixelY(Loc_RS!Top)
        'f_showlc!txt(txt_num).width = pixelx(loc_rs!width)
        'f_showlc!txt(txt_num).Height = pixely(loc_rs!Height)
        
        'f_showlc!txt(txt_num).visible = True
        'f_showlc!txt(txt_num).ZOrder 0
        
      Else
        Status_txt_num = Status_txt_num + 1
        Load F_ShowLc!stxt(Status_txt_num)
        F_ShowLc!stxt(Status_txt_num).Caption = Trim(Loc_RS!prompt)
        F_ShowLc!stxt(Status_txt_num).ForeColor = Loc_RS!font_color
        On Error GoTo skip2
        F_ShowLc!stxt(Status_txt_num).FontName = Trim(Loc_RS!font_name)
skip2:
        On Error GoTo 0
        F_ShowLc!stxt(Status_txt_num).FontSize = Loc_RS!Font_size
        F_ShowLc!stxt(Status_txt_num).Tag = str(Loc_RS!Font_size)
        F_ShowLc!stxt(Status_txt_num).Left = PixelX(Loc_RS!Left)
        F_ShowLc!stxt(Status_txt_num).Top = PixelY(Loc_RS!Top)
        'f_showlc!stxt(status_txt_num).width = pixelx(loc_rs!width)
        'f_showlc!stxt(status_txt_num).Height = pixely(loc_rs!Height)
        
        'f_showlc!stxt(status_txt_num).visible = True
        'f_showlc!stxt(status_txt_num).ZOrder 0
        
        Status_txt(Status_txt_num).Relation_code = Trim(Loc_RS!Relation_code)
      End If
        Loc_RS.MoveNext
    Loop
  Loc_RS.Close
  Loc_DB.Close

'设置pic的min,max值
Set Comm_Db = Workspaces(0).OpenDatabase("c:\ljxt\comm.mdb")
Set Comm_RS = Comm_Db.OpenRecordset("send_table", dbOpenTable)
Comm_RS.Index = "mn_idx"
For K = 1 To Pic_num
    If Dev_pic(K).Style = 7 Then
        Comm_RS.Seek "=", Dev_pic(K).Machine, Trim(Dev_pic(K).Name)
'        Comm_rs.FindFirst "mathine=" + str(Dev_pic(k).Machine) + " and trim(name)='" + Trim(Dev_pic(k).Name) + "'"
         If Not Comm_RS.NoMatch Then
             Dev_pic(K).Min = Comm_RS!Min
             Dev_pic(K).Max_size = Comm_RS!Max
         End If
    End If
Next K
Comm_RS.Close
Comm_Db.Close
  
  For i = 1 To Txt_num
    F_ShowLc!txt(i).Visible = True
    F_ShowLc!txt(i).ZOrder 0
  Next i
'  For i = 1 To status_txt_num
'    f_showlc!stxt(i).visible = True
'    f_showlc!stxt(i).ZOrder 0
'  Next i
End Sub


Sub main()
'Set db = OpenDatabase(App.Path + "\" + "allmdb.mdb")
Set db = OpenDatabase("c:\ljxt\tkbmp\allmdb.mdb")
Set rs = db.OpenRecordset("auth_table", dbOpenDynaset)
Combo1.Clear
End Sub
Private Sub Combo1_Change()
  rs.FindFirst "name='" + Trim(Combo1.Text) + "'"
  If Not rs.NoMatch Then
      Text1.Text = Trim(rs!author_name)
      Text2.Text = Trim(rs!descript)
  Else
      Text1.Text = "": Text2.Text = ""
  End If
End Sub

Private Sub Combo1_Click()
  rs.FindFirst "name='" + Trim(Combo1.Text) + "'"
  If Not rs.NoMatch Then
      Text1.Text = Trim(rs!author_name)
      Text2.Text = Trim(rs!descript)
  Else
      Text1.Text = "": Text2.Text = ""
  End If
End Sub


Private Sub Form_Activate()
'Set db = OpenDatabase(App.Path + "\" + "allmdb.mdb")
Set db = OpenDatabase("c:\ljxt\tkbmp\allmdb.mdb")
Set rs = db.OpenRecordset("auth_table", dbOpenDynaset)
Call main
Do Until rs.EOF
  Combo1.AddItem Trim(rs!Name)
  rs.MoveNext
Loop
Text1.Text = "": Text2.Text = ""
End Sub

Private Sub SSCommand1_Click()
Dim Answer As Integer
  rs.FindFirst "name='" + Trim(Combo1.Text) + "'"
  If Not rs.NoMatch Then
        If F_ShowLc.Visible = True Then
            Unload F_ShowLc
            Load F_ShowLc
        End If
        f_read.Hide
        F_ShowLc.BackColor = rs!back_color
        F_ShowLc.Show
        F_ShowLc.Left = 0
        F_ShowLc.Top = 0
        F_ShowLc.Width = rs!Width
        F_ShowLc.Height = rs!Height
        Old_Window_Width = F_ShowLc.Width
        Old_Window_Height = F_ShowLc.Height
        F_ShowLc.ScaleWidth = PixelX(F_ShowLc.Width)
        F_ShowLc.ScaleHeight = PixelX(F_ShowLc.Height)
        Old_Window_ScaleWidth = F_ShowLc.ScaleWidth
        Old_Window_ScaleHeight = F_ShowLc.ScaleHeight
        Gqg.Color = rs!q_color: Gqg.Width = rs!q_width
        Gyg.Color = rs!y_color: Gyg.Width = rs!y_width
        Glg.Color = rs!j_color: Glg.Width = rs!j_width
        
        Cur_filename = Trim(rs!Name)
        MAIN_MDI.Caption = "炼胶系统-正在演示流程图：(" + Trim(Cur_filename) + ")"
        F_ShowLc.Timer1.Enabled = False
        Call load_from_mdb(Trim(rs!pathname))
        F_ShowLc.Timer1.Enabled = True
  Else
        MsgBox "未找到图库：" + Trim(Combo1.Text) + "！", vbInformation, "提示"
  End If
End Sub


Private Sub SSCommand2_Click()
   rs.Close
   db.Close
   f_read.Hide
End Sub



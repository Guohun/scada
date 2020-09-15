VERSION 4.00
Begin VB.Form f_open 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打开图库"
   ClientHeight    =   4344
   ClientLeft      =   996
   ClientTop       =   1488
   ClientWidth     =   5772
   Height          =   4728
   Left            =   948
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4344
   ScaleWidth      =   5772
   ShowInTaskbar   =   0   'False
   Top             =   1152
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
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   9.6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox Text2 
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   9.6
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
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
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   9.6
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1440
         Width           =   2532
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   9.6
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2532
      End
      Begin VB.Label Label3 
         Caption         =   "说    明："
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "作    者："
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "图库名："
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
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
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
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
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
End
Attribute VB_Name = "f_open"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim db As Database, rs As Recordset

'从数据库取出元件数据
Sub load_from_mdb(mdb_name As String)
Dim loc_db As Database, loc_rs, sl_rs As Recordset
Dim i, j, zxy_index As Integer
Dim head As String
Dim temp_dot As xy_type
If Trim(mdb_name) = "" Or Dir$(Trim(mdb_name)) = Null Then
  Exit Sub
End If
Set loc_db = Workspaces(0).OpenDatabase(Trim(mdb_name))
Set loc_rs = loc_db.OpenRecordset("grap_table", dbOpenDynaset)
Set rl_rs = loc_db.OpenRecordset("relation_table", dbOpenDynaset)
Set sl_rs = loc_db.OpenRecordset("sub_line_table", dbOpenDynaset)

'读取图象
pic_num = 0
selected_pic = -1
loc_rs.FindFirst "trim(code)>='O00000' and trim(code)<='O99999'"
Do Until loc_rs.NoMatch
    pic_num = pic_num + 1
    dev_pic(pic_num).code = Trim(loc_rs!code)
    dev_pic(pic_num).Name = Trim(loc_rs!Name)
    dev_pic(pic_num).Style = loc_rs!Style
    dev_pic(pic_num).position = loc_rs!position

    Load f_paint!p(pic_num)
    'need code
    Call f_paint.load_pic(pic_num, dev_pic(pic_num).Style)
    f_paint!p(pic_num).Left = loc_rs!Left
    f_paint!p(pic_num).Top = loc_rs!Top
    f_paint!p(pic_num).Width = loc_rs!Width
    f_paint!p(pic_num).Height = loc_rs!Height
    
    f_paint!p(pic_num).Visible = True
    f_paint!p(pic_num).ZOrder 0
    
    o_num = o_num + 1
    loc_rs.FindNext "trim(code)>='O00000' and trim(code)<='O99999'"
Loop

'------------------------------------------------
'读取折线(relation_table,sub_line_table)
zline_num = 0
selected_zline = -1
line_num = 0
j_num = 0: y_num = 0: q_num = 0
'供胶管（料）
    '由relation_table查找管道
    loc_rs.FindFirst "trim(code)>='J00000' and trim(code)<='J99999'"
    Do Until loc_rs.NoMatch
        dev_zline(zline_num).code = Trim(loc_rs!code)
        dev_zline(zline_num).Name = Trim(loc_rs!Name)
        dev_zline(zline_num).Style = loc_rs!Style
        
        'start_pic_index,   end_pic_index
        rl_rs.FindFirst "code='" + Trim(dev_zline(zline_num).code) + "'"
        If Not rl_rs.NoMatch Then
            For i = 1 To pic_num
                If Trim(dev_pic(i).code) = Trim(rl_rs!last_code) Then
                    dev_zline(zline_num).start_pic_index = i
                    Exit For
                End If
            Next i
            For i = 1 To pic_num
                If Trim(dev_pic(i).code) = Trim(rl_rs!next_code) Then
                    dev_zline(zline_num).end_pic_index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        zxy_index = 0
        sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=0"
        If Not sl_rs.NoMatch Then
            temp_dot.Left = sl_rs!Left
            temp_dot.Top = sl_rs!Top
        End If
        dev_zline(zline_num).first_line_index = line_num + 1
        
        zxy_index = zxy_index + 1
        sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=" + Str(zxy_index)
        Do Until sl_rs.NoMatch
            line_num = line_num + 1
            Load f_paint!l(line_num)
            f_paint!l(line_num).BorderWidth = glg.Width
            f_paint!l(line_num).BorderColor = glg.Color
            f_paint!l(line_num).X1 = temp_dot.Left
            f_paint!l(line_num).Y1 = temp_dot.Top
            f_paint!l(line_num).X2 = sl_rs!Left
            f_paint!l(line_num).Y2 = sl_rs!Top
            f_paint!l(line_num).Visible = True
            
            temp_dot.Left = sl_rs!Left
            temp_dot.Top = sl_rs!Top
            zxy_index = zxy_index + 1
            sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=" + Str(zxy_index)
        Loop
        dev_zline(zline_num).last_line_index = line_num
        
        j_num = j_num + 1
        zline_num = zline_num + 1
        
        loc_rs.FindNext "trim(code)>='J00000' and trim(code)<='J99999'"
    Loop
  
'供汽管
    '由relation_table查找管道
    loc_rs.FindFirst "trim(code)>='Q00000' and trim(code)<='Q99999'"
    Do Until loc_rs.NoMatch
        dev_zline(zline_num).code = Trim(loc_rs!code)
        dev_zline(zline_num).Name = Trim(loc_rs!Name)
        dev_zline(zline_num).Style = loc_rs!Style
        
        'start_pic_index,   end_pic_index
        rl_rs.FindFirst "code='" + Trim(dev_zline(zline_num).code) + "'"
        If Not rl_rs.NoMatch Then
            For i = 1 To pic_num
                If Trim(dev_pic(i).code) = Trim(rl_rs!last_code) Then
                    dev_zline(zline_num).start_pic_index = i
                    Exit For
                End If
            Next i
            For i = 1 To pic_num
                If Trim(dev_pic(i).code) = Trim(rl_rs!next_code) Then
                    dev_zline(zline_num).end_pic_index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        zxy_index = 0
        sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=0"
        If Not sl_rs.NoMatch Then
            temp_dot.Left = sl_rs!Left
            temp_dot.Top = sl_rs!Top
        End If
        dev_zline(zline_num).first_line_index = line_num + 1
        
        zxy_index = zxy_index + 1
        sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=" + Str(zxy_index)
        Do Until sl_rs.NoMatch
            line_num = line_num + 1
            Load f_paint!l(line_num)
            f_paint!l(line_num).BorderWidth = gqg.Width
            f_paint!l(line_num).BorderColor = gqg.Color
            f_paint!l(line_num).X1 = temp_dot.Left
            f_paint!l(line_num).Y1 = temp_dot.Top
            f_paint!l(line_num).X2 = sl_rs!Left
            f_paint!l(line_num).Y2 = sl_rs!Top
            f_paint!l(line_num).Visible = True
            
            temp_dot.Left = sl_rs!Left
            temp_dot.Top = sl_rs!Top
            zxy_index = zxy_index + 1
            sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=" + Str(zxy_index)
        Loop
        dev_zline(zline_num).last_line_index = line_num
        
        q_num = q_num + 1
        zline_num = zline_num + 1
        
        loc_rs.FindNext "trim(code)>='Q00000' and trim(code)<='Q99999'"
    Loop
  
'供油管
    '由relation_table查找管道
    loc_rs.FindFirst "trim(code)>='Y00000' and trim(code)<='Y99999'"
    Do Until loc_rs.NoMatch
        dev_zline(zline_num).code = Trim(loc_rs!code)
        dev_zline(zline_num).Name = Trim(loc_rs!Name)
        dev_zline(zline_num).Style = loc_rs!Style
        
        'start_pic_index,   end_pic_index
        rl_rs.FindFirst "code='" + Trim(dev_zline(zline_num).code) + "'"
        If Not rl_rs.NoMatch Then
            For i = 1 To pic_num
                If Trim(dev_pic(i).code) = Trim(rl_rs!last_code) Then
                    dev_zline(zline_num).start_pic_index = i
                    Exit For
                End If
            Next i
            For i = 1 To pic_num
                If Trim(dev_pic(i).code) = Trim(rl_rs!next_code) Then
                    dev_zline(zline_num).end_pic_index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        zxy_index = 0
        sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=0"
        If Not sl_rs.NoMatch Then
            temp_dot.Left = sl_rs!Left
            temp_dot.Top = sl_rs!Top
        End If
        dev_zline(zline_num).first_line_index = line_num + 1
        
        zxy_index = zxy_index + 1
        sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=" + Str(zxy_index)
        Do Until sl_rs.NoMatch
            line_num = line_num + 1
            Load f_paint!l(line_num)
            f_paint!l(line_num).BorderWidth = gyg.Width
            f_paint!l(line_num).BorderColor = gyg.Color
            f_paint!l(line_num).X1 = temp_dot.Left
            f_paint!l(line_num).Y1 = temp_dot.Top
            f_paint!l(line_num).X2 = sl_rs!Left
            f_paint!l(line_num).Y2 = sl_rs!Top
            f_paint!l(line_num).Visible = True
            
            temp_dot.Left = sl_rs!Left
            temp_dot.Top = sl_rs!Top
            zxy_index = zxy_index + 1
            sl_rs.FindFirst "trim(code)='" + Trim(dev_zline(zline_num).code) + "' and index=" + Str(zxy_index)
        Loop
        dev_zline(zline_num).last_line_index = line_num
        
        y_num = y_num + 1
        zline_num = zline_num + 1
        
        loc_rs.FindNext "trim(code)>='Y00000' and trim(code)<='Y99999'"
    Loop
    sl_rs.Close
    rl_rs.Close
'------------------------------------------------
'读取信号灯(grap_table)
xhd_num = 0
selected_xhd = -1
loc_rs.FindFirst "trim(code)>='X00000' and trim(code)<='X99999'"
Do Until loc_rs.NoMatch
    xhd_num = xhd_num + 1
    dev_xhd(xhd_num).code = Trim(loc_rs!code)
    dev_xhd(xhd_num).Name = Trim(loc_rs!Name)

    Load f_paint!xhd(xhd_num)
    'need code
    f_paint!xhd(xhd_num).Left = loc_rs!Left
    f_paint!xhd(xhd_num).Top = loc_rs!Top
    f_paint!xhd(xhd_num).Width = loc_rs!Width
    f_paint!xhd(xhd_num).Height = loc_rs!Height
    
    f_paint!xhd(xhd_num).Visible = True
    f_paint!xhd(xhd_num).ZOrder 0
    
    x_num = x_num + 1
    loc_rs.FindNext "trim(code)>='X00000' and trim(code)<='X99999'"
Loop
loc_rs.Close
    
'读取文字(prompt_table)
    txt_num = 0
    Set loc_rs = loc_db.OpenRecordset("prompt_table", dbOpenDynaset)
    Do Until loc_rs.EOF
        txt_num = txt_num + 1
        Load f_paint!txt(txt_num)
        f_paint!txt(txt_num).Caption = Trim(loc_rs.Prompt)
        f_paint!txt(txt_num).ForeColor = loc_rs!font_color
        f_paint!txt(txt_num).FontName = Trim(loc_rs!font_name)
        f_paint!txt(txt_num).FontSize = loc_rs!font_size
        f_paint!txt(txt_num).Left = loc_rs!Left
        f_paint!txt(txt_num).Top = loc_rs!Top
        'f_paint!txt(txt_num).width = loc_rs!width
        'f_paint!txt(txt_num).Height = loc_rs!Height
        f_paint!txt(txt_num).Visible = True
        f_paint!txt(txt_num).ZOrder 0
          
        dev_txt(txt_num).Style = loc_rs!Style
        dev_txt(txt_num).relation_code = Trim(loc_rs!relation_code)
        loc_rs.MoveNext
    Loop
  loc_rs.Close
  loc_db.Close
End Sub

Sub main()
Set db = OpenDatabase(App.Path + "\" + "allmdb.mdb")
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
Set db = OpenDatabase(App.Path + "\" + "allmdb.mdb")
Set rs = db.OpenRecordset("auth_table", dbOpenDynaset)
Call main
Do Until rs.EOF
  Combo1.AddItem Trim(rs!Name)
  rs.MoveNext
Loop
Text1.Text = "": Text2.Text = ""
End Sub

Private Sub SSCommand1_Click()
Dim answer As Integer
  rs.FindFirst "name='" + Trim(Combo1.Text) + "'"
  If Not rs.NoMatch Then
     If f_paint.Visible = True Then
        answer = MsgBox("是否保存图库：(" + Trim(cur_filename) + ")", vbQuestion + vbYesNoCancel, "警告")
        If answer = vbYes Then
           f_save.Show vbModal
        End If
        If answer = vbCancel Then
            Exit Sub
        End If
    End If
        Unload f_paint
        Call f_main.main
        f_open.Hide
        Load f_paint
        f_paint.BackColor = rs!back_color
        f_paint.Show
        f_paint.Width = rs!Width
        f_paint.Height = rs!Height
        gqg.Color = rs!q_color: gqg.Width = rs!q_width
        gyg.Color = rs!y_color: gyg.Width = rs!y_width
        glg.Color = rs!j_color: glg.Width = rs!j_width
        
        cur_filename = Trim(rs!Name)
        f_main.Caption = "图形编辑器。正在编辑图库：(" + Trim(cur_filename) + ")"
        Call load_from_mdb(Trim(rs!pathname))
  Else
        MsgBox "未找到图库：" + Trim(Combo1.Text) + "！", vbInformation, "提示"
  End If
End Sub

Private Sub SSCommand2_Click()
   rs.Close
   db.Close
   f_open.Hide
End Sub

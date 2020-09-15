VERSION 5.00
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form F_Size 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "缩放窗口"
   ClientHeight    =   2616
   ClientLeft      =   2196
   ClientTop       =   2556
   ClientWidth     =   4212
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2616
   ScaleWidth      =   4212
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   1212
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2652
      _Version        =   65536
      _ExtentX        =   4678
      _ExtentY        =   2138
      _StockProps     =   14
      Caption         =   "设置高宽"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Begin VB.TextBox Text2 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   372
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   732
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   372
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   732
      End
      Begin Spin.SpinButton SpinButton2 
         Height          =   372
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Width           =   252
         _Version        =   65536
         _ExtentX        =   445
         _ExtentY        =   656
         _StockProps     =   73
      End
      Begin Spin.SpinButton SpinButton1 
         Height          =   372
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   252
         _Version        =   65536
         _ExtentX        =   445
         _ExtentY        =   656
         _StockProps     =   73
      End
      Begin VB.Label Label2 
         Caption         =   "高度倍数："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "宽度倍数："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   852
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   972
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2652
      _Version        =   65536
      _ExtentX        =   4678
      _ExtentY        =   1715
      _StockProps     =   14
      Caption         =   "设置"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Begin VB.CheckBox Check1 
         Caption         =   "缩放时改变状态文字的尺寸"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   492
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   1092
      _Version        =   65536
      _ExtentX        =   1926
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "退    出"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   1092
      _Version        =   65536
      _ExtentX        =   1926
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "应    用"
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
Attribute VB_Name = "F_Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：放大或缩小整个演示窗口

Public Function Set_Window_Scale(Scale_X As Double, Scale_Y As Double, sf_with_stxt As Integer)
Dim i As Integer
Dim Tmp_ScaleWidth, Tmp_ScaleHeight As Long
Dim SX, SY As Double
    F_ShowLc.Left = 0: F_ShowLc.Top = 0
    F_ShowLc.Width = CDbl(Old_Window_Width) * Scale_X
    F_ShowLc.Height = CDbl(Old_Window_Height) * Scale_Y
    Tmp_ScaleWidth = F_ShowLc.ScaleWidth
    Tmp_ScaleHeight = F_ShowLc.ScaleHeight
    F_ShowLc.ScaleWidth = Old_Window_ScaleWidth
    F_ShowLc.ScaleHeight = Old_Window_ScaleHeight
    
    For i = 1 To Pic_num
'        Dev_pic(i).Left = CLng(CDbl(Dev_pic(i).Old_Left) * Scale_X)
'        Dev_pic(i).Top = CLng(CDbl(Dev_pic(i).Old_Top) * Scale_Y)
'        Dev_pic(i).Width = CLng(CDbl(Dev_pic(i).Old_Width) * Scale_X)
'        Dev_pic(i).Height = CLng(CDbl(Dev_pic(i).Old_Height) * Scale_Y)
 
 '       readdb_draw_pic (i)
    Next i
    
    For i = 1 To Xhd_num
        Dev_XHD(i).Left = CDbl(Dev_XHD(i).Old_Left) * Scale_X
        Dev_XHD(i).Top = CDbl(Dev_XHD(i).Old_Top) * Scale_Y
        Dev_XHD(i).Width = CDbl(Dev_XHD(i).Old_Width) * Scale_X
        Dev_XHD(i).Height = CDbl(Dev_XHD(i).Old_Height) * Scale_Y
    Next i
    
    For i = 1 To Txt_num
        F_ShowLc!txt(i).Visible = False
        Tmp_X = F_ShowLc!txt(i).Left:      Tmp_Y = F_ShowLc!txt(i).Top
        F_ShowLc!txt(i).FontSize = Val(F_ShowLc!txt(i).Tag) * IIf(Scale_X > Scale_Y, Scale_X, Scale_Y)
        F_ShowLc!txt(i).Left = CDbl(Tmp_X) * Scale_X / Window_Scale_X
        F_ShowLc!txt(i).Top = CDbl(Tmp_Y) * Scale_Y / Window_Scale_Y
        F_ShowLc!txt(i).Visible = True
    Next i
    
    For i = 1 To Status_txt_num
        F_ShowLc!stxt(i).Visible = False
        Tmp_X = F_ShowLc!stxt(i).Left:      Tmp_Y = F_ShowLc!stxt(i).Top
        If sf_with_stxt = 1 Then
            F_ShowLc!stxt(i).FontSize = Val(F_ShowLc!stxt(i).Tag) * IIf(Scale_X > Scale_Y, Scale_X, Scale_Y)
        Else
            F_ShowLc!stxt(i).FontSize = Val(F_ShowLc!stxt(i).Tag)
        End If
        F_ShowLc!stxt(i).Left = CDbl(Tmp_X) * Scale_X / Window_Scale_X
        F_ShowLc!stxt(i).Top = CDbl(Tmp_Y) * Scale_Y / Window_Scale_Y
'        f_showlc!stxt(i).visible = True
    Next i
    
'    For i = 1 To Kj_num
'        F_ShowLc.KJ(i).Left = CDbl(F_ShowLc.KJ(i).Left) * Scale_X / Window_Scale_X
'        F_ShowLc.KJ(i).Top = CDbl(F_ShowLc.KJ(i).Top) * Scale_Y / Window_Scale_Y
'        F_ShowLc.KJ(i).Width = CDbl(F_ShowLc.KJ(i).Width) * Scale_X / Window_Scale_X
'        F_ShowLc.KJ(i).Height = CDbl(F_ShowLc.KJ(i).Height) * Scale_Y / Window_Scale_Y
'    Next i

'    f_showlc.ScaleWidth = old_window_scalewidth / window_scale_x
'    f_showlc.ScaleHeight = old_window_scaleheight / window_scale_y
    Window_Scale_X = Scale_X
    Window_Scale_Y = Scale_Y
    SF_With_Status_Txt = sf_with_stxt
    F_ShowLc.Refresh
End Function


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
Text1.Text = "1.00"
Text2.Text = "1.00"
End Sub


Private Sub SpinButton1_SpinDown()
If Trim(Text1.Text) = "0.30" Then
    Exit Sub
End If
Text1.Text = Format$(Val(Text1.Text) - 0.05, "0.00")
End Sub


Private Sub SpinButton1_SpinUp()
If Trim(Text1.Text) = "5.00" Then
    Exit Sub
End If
Text1.Text = Format$(Val(Text1.Text) + 0.05, "0.00")
End Sub


Private Sub SpinButton2_SpinDown()
If Trim(Text2.Text) = "0.30" Then
    Exit Sub
End If
Text2.Text = Format$(Val(Text2.Text) - 0.05, "0.00")
End Sub

Private Sub SpinButton2_SpinUp()
If Trim(Text2.Text) = "5.00" Then
    Exit Sub
End If
Text2.Text = Format$(Val(Text2.Text) + 0.05, "0.00")
End Sub


Private Sub SSCommand1_Click()
    If Not Val_Comp(Val(Text1.Text), 0.1, 5#) Then
        MsgBox "X轴的缩放比例超出范围(0.10 -- 5.00)!", 48, "错误信息"
        Text1.SetFocus
        Exit Sub
    End If
    If Not Val_Comp(Val(Text2.Text), 0.1, 5#) Then
        MsgBox "Y轴的缩放比例超出范围(0.10 -- 5.00)!", 48, "错误信息"
        Text2.SetFocus
        Exit Sub
    End If
    Set_Window_Scale CDbl(Text1.Text), CDbl(Text2.Text), Check1.Value
End Sub

Private Sub SSCommand2_Click()
    F_Size.Hide
End Sub


Private Sub Text1_LostFocus()
    If Not Val_Comp(Val(Text1.Text), 0.1, 5#) Then
        MsgBox "X轴的缩放比例超出范围(0.10 -- 5.00)!", 48, "错误信息"
        Text1.SetFocus
        'Text1.Text = IIf(val(Text1.Text) > 5#, "5.00", "0.10")
    End If
End Sub


Private Sub Text2_LostFocus()
    If Not Val_Comp(Val(Text2.Text), 0.1, 5#) Then
        MsgBox "Y轴的缩放比例超出范围(0.10 -- 5.00)!", 48, "错误信息"
        Text2.SetFocus
        'Text2.Text = IIf(val(Text2.Text) > 5#, "5.00", "0.10")
    End If
End Sub



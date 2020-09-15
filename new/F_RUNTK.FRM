VERSION 4.00
Begin VB.MDIForm f_runtk 
   BackColor       =   &H00808080&
   Caption         =   "炼胶系统生产流程图"
   ClientHeight    =   5640
   ClientLeft      =   12
   ClientTop       =   1548
   ClientWidth     =   9576
   Height          =   6312
   Icon            =   "f_runtk.frx":0000
   Left            =   -36
   LinkTopic       =   "MDIForm1"
   Top             =   924
   Width           =   9672
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   9576
      _Version        =   65536
      _ExtentX        =   16891
      _ExtentY        =   847
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      Begin VB.Label infor 
         Caption         =   "信息"
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9252
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Menu m_file 
      Caption         =   "文件"
      Begin VB.Menu m_f_open 
         Caption         =   "打开图库"
      End
      Begin VB.Menu m_f_close 
         Caption         =   "关闭图库"
      End
   End
   Begin VB.Menu m_option 
      Caption         =   "设置"
      Begin VB.Menu m_o_size 
         Caption         =   "窗口缩放"
      End
      Begin VB.Menu m_o_gd 
         Caption         =   "管道内径及颜色"
      End
      Begin VB.Menu m_o_refresh 
         Caption         =   "刷新窗口"
      End
      Begin VB.Menu m_o_space 
         Caption         =   "-"
      End
      Begin VB.Menu m_o_error 
         Caption         =   "显示错误窗口"
      End
      Begin VB.Menu m_o_infor 
         Caption         =   "隐藏信息条 "
      End
   End
   Begin VB.Menu m_quit 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "f_runtk"
Attribute VB_Creatable = False
Attribute VB_Exposed = False









Private Sub m_f_close_Click()
If F_showlc.Visible = True Then
    Unload F_showlc
End If
End Sub


Private Sub m_f_open_Click()
    f_read.Show vbModal
End Sub


Private Sub m_o_error_Click()
    f_error.Show
    f_error.ZOrder 0
End Sub

Private Sub m_o_gd_Click()
    f_gd_lc.Show vbModal
End Sub

Private Sub m_o_infor_Click()
    If f_runtk!SSPanel1.Visible = True Then
        m_o_infor.Caption = "显示信息条"
        f_runtk!SSPanel1.Visible = False
    Else
        m_o_infor.Caption = "隐藏信息条"
        f_runtk!SSPanel1.Visible = True
    End If
End Sub

Private Sub m_o_refresh_Click()
    F_showlc.Refresh
End Sub

Private Sub m_o_size_Click()
    f_size.Show vbModal
End Sub

Private Sub m_quit_Click()
    End
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    Unload Me
End Sub



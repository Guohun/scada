VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form F_ShowLc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5988
   ClientLeft      =   12
   ClientTop       =   636
   ClientWidth     =   9576
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   Begin VB.Data ComData 
      Caption         =   "Component Data"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\tkbmp\AllMdb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Component_Table"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2652
   End
   Begin Threed.SSPanel StatusPanel 
      Height          =   1572
      Left            =   960
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   5172
      _Version        =   65536
      _ExtentX        =   9123
      _ExtentY        =   2773
      _StockProps     =   15
      ForeColor       =   16777215
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Begin VB.Data ErrorHistoryData 
         Caption         =   "ErrorHistoryData"
         Connect         =   "Access"
         DatabaseName    =   "C:\ljxt\Ljxt.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   492
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   "Error_History_Table"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2172
      End
      Begin MSGrid.Grid ErrorGrid 
         Height          =   732
         Left            =   120
         TabIndex        =   0
         Top             =   840
         Width           =   4932
         _Version        =   65536
         _ExtentX        =   8700
         _ExtentY        =   1291
         _StockProps     =   77
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Cols            =   5
         ScrollBars      =   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   9
         Left            =   3480
         TabIndex        =   25
         Top             =   600
         Width           =   492
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   9
         Left            =   4080
         TabIndex        =   24
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "压力:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   8
         Left            =   2040
         TabIndex        =   23
         Top             =   600
         Width           =   492
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   8
         Left            =   2640
         TabIndex        =   22
         Top             =   600
         Width           =   612
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "管道压力:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   900
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   7
         Left            =   1200
         TabIndex        =   20
         Top             =   600
         Width           =   612
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   4
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "段时间:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   696
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   6
         Left            =   4560
         TabIndex        =   2
         Top             =   360
         Width           =   492
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   492
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   5
         Left            =   4560
         TabIndex        =   4
         Top             =   120
         Width           =   492
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   2
         Left            =   3000
         TabIndex        =   6
         Top             =   120
         Width           =   492
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   252
      End
      Begin VB.Label StatusVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Top             =   120
         Width           =   372
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "车数:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   492
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "运行车号:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "段号:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Width           =   492
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "总时间:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   4
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   696
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "温度:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   5
         Left            =   3840
         TabIndex        =   13
         Top             =   120
         Width           =   492
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "能量:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   6
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   492
      End
   End
   Begin VB.PictureBox PicDemo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DataField       =   "ComPicUp"
      DataSource      =   "ComData"
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   972
      Left            =   1200
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.PictureBox PicMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DataField       =   "MaskPicUp"
      DataSource      =   "ComData"
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   972
      Left            =   240
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   600
      Top             =   3720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3240
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327680
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   3240
   End
   Begin VB.Data ErrorData 
      Caption         =   "ErrorData"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   1680
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "error_num"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.Shape xhd 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   0
      Left            =   240
      Shape           =   2  'Oval
      Top             =   3960
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label txt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "标注文字"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   1068
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label stxt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "标注状态"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   1080
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Menu m_window 
      Caption         =   "窗口(&W)"
      Begin VB.Menu m_o_StopSpeaking 
         Caption         =   "停止报警"
         Shortcut        =   {F2}
      End
      Begin VB.Menu m_o_errorhelp 
         Caption         =   "错误帮助"
         Shortcut        =   {F3}
      End
      Begin VB.Menu m_o_errhis 
         Caption         =   "错误历史"
         Shortcut        =   ^H
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu m_w_liucheng 
         Caption         =   "关闭演示窗口"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu m_option 
      Caption         =   "设置(&O)"
      Begin VB.Menu m_w_main_tk 
         Caption         =   "选择系统图"
         Shortcut        =   ^S
      End
      Begin VB.Menu m_o_zoom 
         Caption         =   "缩放窗口"
         Enabled         =   0   'False
         Shortcut        =   ^Z
         Visible         =   0   'False
      End
      Begin VB.Menu m_o_guandao 
         Caption         =   "管道属性"
         Shortcut        =   ^G
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu AddErr 
         Caption         =   "增加错误"
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu DelErr 
         Caption         =   "删除错误"
         Shortcut        =   ^D
         Visible         =   0   'False
      End
      Begin VB.Menu m_o_tk_machine 
         Caption         =   "图库与机组对应关系"
      End
      Begin VB.Menu sp4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu m_o_test 
         Caption         =   "流体运动仿真算法测试"
         Shortcut        =   ^T
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "F_ShowLc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：实时对现场生产过程进行仿真演示

Dim temp As Integer                         '临时变量
Dim In_Timer1, In_Timer2 As Boolean         '进入标志（定时器用）

'**************************************************************************
'*  函数功能：将当前产生的错误信息存入数据库中                              *
'**************************************************************************
Public Function Save_Last_Error_To_DB()
Dim i As Integer
Dim MaxGridErrorNum, MaxDBErrorNum, DBDelErrorNum As Integer
MaxGridErrorNum = 50                                        '演示窗口中能浏览的最近产生的错误信息的个数
MaxDBErrorNum = 200                                         '数据库中允许存放的最多错误记录数
DBDelErrorNum = 10                                              '系统一次删除错误信息的记录数
'将最新的错误存入库中
With ErrorHistoryData.Recordset
    .AddNew
    ErrorGrid.col = 1:            !Error_Code = CInt(ErrorGrid.Text)
    ErrorGrid.col = 2:              !Error_Prompt = Trim(ErrorGrid.Text)
    ErrorGrid.col = 3:              !Error_time = Trim(ErrorGrid.Text)
    ErrorGrid.col = 4:              !Error_date = Trim(ErrorGrid.Text)
    .Update
End With
'清库
ErrorHistoryData.Recordset.MoveLast
If ErrorHistoryData.Recordset.RecordCount > MaxDBErrorNum Then
    ErrorHistoryData.Recordset.MoveFirst
    For i = 1 To DBDelErrorNum
        ErrorHistoryData.Recordset.Delete
        ErrorHistoryData.Refresh
    Next i
    ErrorHistoryData.Recordset.MoveFirst
End If

'清错误历史网格控件
With ErrorGrid
    If .Rows > MaxGridErrorNum + 1 Then
        .TopRow = 2
        .RemoveItem 1
        .TopRow = .Rows - 1
    End If
End With
End Function

'**************************************************************************
'*  函数功能：显示报警信息及声音（定时器用）                                *
'**************************************************************************
Public Function Run_Error()
On Error Resume Next
Dim i, j As Integer
Dim SndFileName As String
Static Tmp, error_num, Speak_Times As Integer
Static Is_Old_Error, Have_First_Error As Boolean
Dim Found_Error As Integer      '0 -- No found any error
                                '1 -- Found old error
                                '2 -- Found new error
Tmp = get_error(Cur_Errors)     '获取错误数据
'将新错误装入错误列表
Found_Error = 0
For i = 0 To Cur_Errors.Length - 1
     Is_Old_Error = False
    For j = 0 To Last_Errors.Length - 1
        If Cur_Errors.Code(i) = Last_Errors.Code(j) Then
            Is_Old_Error = True
            Exit For
        End If
    Next j
    If Not Is_Old_Error Then   '有新错误
        With ErrorGrid
            error_num = error_num + 1
            '清除当前标志
            If Have_First_Error Then
                .row = .Rows - 1:        .col = 0
                .Text = str(error_num - 1)
                .Rows = .Rows + 1
            Else
                Have_First_Error = True
            End If
            .row = .Rows - 1
            .col = 0:                .Text = "*" + str(error_num)
            .col = 1:                .Text = Left(Trim(CStr(Cur_Errors.Code(i))), 4)
            .col = 2:
        End With  'change by zgh    1998 -3 -9
            ErrorData.Recordset.index = "error_idx"
            ErrorData.Recordset.Seek "=", CStr(Cur_Errors.Code(i))
            '在库中找到错误信息
        
            If Not ErrorData.Recordset.NoMatch Then
                ErrorGrid.Text = Trim(ErrorData.Recordset!china_prompt)
            Else
                ErrorGrid.Text = Trim(Cur_Errors.Prompt(i))
            End If
        With ErrorGrid      'change by zgh 1998 3 8
            .col = 3:               .Text = Left(time$(), 5)
            .col = 4:               .Text = Right(Date$, 8)
            .col = 1
            .TopRow = .row
        End With
        Cur_Speaking_Error_Code = Cur_Errors.Code(i)
        Found_Error = 2
        Speak_Times = 0
        Call Save_Last_Error_To_DB      '保存错误信息到库中
    Else        '有旧错误
        Cur_Speaking_Error_Code = Cur_Errors.Code(i)
        Found_Error = 1
    End If      '有新错误
Next i
If Found_Error = 0 Then
    Cur_Speaking_Error_Code = -1
    Speak_Times = 0
End If
For i = 0 To Cur_Errors.Length - 1
    Last_Errors.Code(i) = Cur_Errors.Code(i)
    Last_Errors.Prompt(i) = Cur_Errors.Prompt(i)
Next i
Last_Errors.Length = Cur_Errors.Length
    
'Tmp = DoEvents
'播放声音
If Speak_Times < 3 Then     '一个错误报警3次
    SndFileName = "c:\ljxt\wave\china\" + Trim(CStr(Cur_Speaking_Error_Code)) + ".WAV"
    If (Dir(SndFileName, vbNormal) <> "") And (Cur_Speaking_Error_Code <> -1) And (Not F_Error_Help.Visible) Then '在库中找到错误信息
        No_Speaking = sndPlaySound(SndFileName, SND_ASYNC + SND_NOSTOP)
    End If
    If No_Speaking Then
        Speak_Times = Speak_Times + 1
    End If
    Last_Speaking_Error_Code = Cur_Speaking_Error_Code
Else    '报警次数超出范围
    For i = 0 To Cur_Errors.Length - 1
        Call kill_error(CLng(Cur_Errors.Code(i)))   '停止报警（清除错误）
    Next i
End If
End Function

'***************************************************************************
'*  函数功能：运行所有密炼机（FORCE＝TRUE 强行显示数据）（定时器用）          *
'***************************************************************************
Public Function Run_LJLS(Force_Show As Boolean)
Dim i As Integer
For i = 1 To Pic_num
    If Dev_PIC(i).Style = 9 Then '密炼机
        Call Run_LJL(i, Force_Show)
    End If
Next i
End Function

'***************************************************************************
'*  函数功能：运行指定密炼机的上顶栓（FORCE＝TRUE 强行显示数据）（密炼机用）  *
'***************************************************************************
Public Function Run_LJL_SDS(LJL_IDX As Integer, index As Integer, Force_Show As Boolean)
On Error GoTo Run_ljl_error
Dim Val_Chg As Boolean
Dim V As Integer        '0--Low         1--High
Dim LJL_W, LJL_H As Single
Dim SDS_L, SDS_T, SDS_W, SDS_H As Single
Dim v1, v2, v3 As Integer
If (LJL(index).SDS_High_IDX = -1) And (LJL(index).SDS_Low_IDX = -1) Then
    Exit Function
End If
LJL_W = 76: LJL_H = 179
SDS_L = 33: SDS_T = 6: SDS_W = 6:  SDS_H = 35
    
If LJL(index).SDS_High_Style = 1 Then 'pic
    v3 = Dev_PIC(LJL(index).SDS_High_IDX).Value
Else
    v3 = Dev_XHD(LJL(index).SDS_High_IDX).Value
End If
If LJL(index).SDS_Mid_Style = 1 Then 'pic
    v2 = Dev_PIC(LJL(index).SDS_Mid_IDX).Value
Else
    v2 = Dev_XHD(LJL(index).SDS_Mid_IDX).Value
End If
If LJL(index).SDS_Low_Style = 1 Then 'pic
    v1 = Dev_PIC(LJL(index).SDS_Low_IDX).Value
Else
    v1 = Dev_XHD(LJL(index).SDS_Low_IDX).Value
End If

V = 0
If v3 = 1 Then
    V = 2
Else
    If v2 = 1 Then
        V = 1
    Else
        V = 0
    End If
End If
LJL(index).SDS_Value = V

If LJL(index).Last_SDS_Status <> LJL(index).SDS_Value Then
    Val_Chg = True
Else
    Val_Chg = False
End If
If Val_Chg Or Force_Show Then
    'High
    Select Case LJL(index).SDS_Value
        Case 2  'high
            Me.Line (Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * SDS_L / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * SDS_T / LJL_H)-(Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * (SDS_L + SDS_W) / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * (SDS_T + SDS_H) / LJL_H), QBColor(12), BF
        Case 1  'mid
            Me.Line (Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * SDS_L / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * SDS_T / LJL_H)-(Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * (SDS_L + SDS_W) / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * (SDS_T + SDS_H) / LJL_H), QBColor(12), BF
            Me.Line (Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * SDS_L / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * SDS_T / LJL_H)-(Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * (SDS_L + SDS_W) / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * (SDS_T + SDS_H) / LJL_H / 2), QBColor(2), BF
        Case 0  'low
            Me.Line (Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * SDS_L / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * SDS_T / LJL_H)-(Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * (SDS_L + SDS_W) / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * (SDS_T + SDS_H) / LJL_H), QBColor(2), BF
    End Select
    LJL(index).Last_SDS_Status = LJL(index).SDS_Value
End If
Exit Function
Run_ljl_error:
           Call write_error(3001, "运行指定密炼机的上顶栓错误")
    
End Function

'***************************************************************************
'*  函数功能：运行指定密炼机的下顶栓（FORCE＝TRUE 强行显示数据）（密炼机用）  *
'***************************************************************************
Public Function Run_LJL_XDS(LJL_IDX As Integer, index As Integer, Force_Show As Boolean)
On Error GoTo Run_ljl_error

Dim Val_Chg As Boolean
Dim V As Integer        '0--Low         1--High
Dim LJL_W, LJL_H As Single
Dim XDS_L, XDS_T, XDS_W, XDS_H As Single
Dim XDS_O_L, XDS_O_T, XDS_O_W, XDS_O_H As Single
If (LJL(index).XDS_OPEN_FM_IDX = -1) And (LJL(index).XDS_CLOSE_FM_IDX = -1) Then
    Exit Function
End If
LJL_W = 76: LJL_H = 179
XDS_L = 3: XDS_T = 85: XDS_W = 66: XDS_H = 7
XDS_O_L = 12 + 2: XDS_O_T = 87 - 2: XDS_O_W = 48 - 4: XDS_O_H = 5 + 4
    
Select Case LJL(index).XDS_OPEN_FM_Style
    Case 0      'Xhd
        V = 0
        If (Dev_XHD(LJL(index).XDS_OPEN_FM_IDX).Value = 1) And (Dev_XHD(LJL(index).XDS_CLOSE_FM_IDX).Value = 0) Then
            V = 1       'Open
        Else
            If (Dev_XHD(LJL(index).XDS_OPEN_FM_IDX).Value = 0) And (Dev_XHD(LJL(index).XDS_CLOSE_FM_IDX).Value = 1) Then
                V = 0   'Close
            End If
        End If
        LJL(index).XDS_Value = V
    Case 1      'Pic
        V = 0
        If (Dev_PIC(LJL(index).XDS_OPEN_FM_IDX).Value = 1) And (Dev_PIC(LJL(index).XDS_CLOSE_FM_IDX).Value = 0) Then
            V = 1       'High
        Else
            If (Dev_PIC(LJL(index).XDS_OPEN_FM_IDX).Value = 0) And (Dev_PIC(LJL(index).XDS_CLOSE_FM_IDX).Value = 1) Then
                V = 0   'Low
            End If
        End If
        LJL(index).XDS_Value = V
End Select

If LJL(index).Last_XDS_Status <> LJL(index).XDS_Value Then
    Val_Chg = True
Else
    Val_Chg = False
End If
If Val_Chg Or Force_Show Then
    Me.FillStyle = 0
    If LJL(index).XDS_Value = 1 Then    'high
        Me.Line (Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * XDS_O_L / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * XDS_O_T / LJL_H)-(Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * (XDS_O_L + XDS_O_W) / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * (XDS_O_T + XDS_O_H) / LJL_H), QBColor(2), BF
    Else
        If LJL(index).XDS_Value = 0 Then 'Low
            Me.Line (Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * XDS_L / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * XDS_T / LJL_H)-(Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * (XDS_L + XDS_W) / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * (XDS_T + XDS_H) / LJL_H), QBColor(1), BF
        End If
    End If
    Me.FillStyle = 1
    Me.Line (Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * XDS_L / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * XDS_T / LJL_H)-(Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * (XDS_L + XDS_W) / LJL_W, Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * (XDS_T + XDS_H) / LJL_H), QBColor(0), B
    LJL(index).Last_XDS_Status = LJL(index).XDS_Value
End If
Exit Function
Run_ljl_error:
           Call write_error(3001, "运行指定密炼机的下顶栓错误")
End Function

'***************************************************************************
'*  函数功能：运行指定密炼机的加料门（FORCE＝TRUE 强行显示数据）（密炼机用）  *
'***************************************************************************
Public Function Run_LJL_JLM(LJL_IDX As Integer, index As Integer, Force_Show As Boolean)
Dim Val_Chg As Boolean
Dim i, V As Integer       '0--Low         1--High
Dim LJL_W, LJL_H As Single
Dim Zxy(0 To 3) As POINTAPI
If (LJL(index).JLM_OPEN_FM_IDX = -1) And (LJL(index).JLM_CLOSE_FM_IDX = -1) Then
    Exit Function
End If

LJL_W = 76: LJL_H = 179
Zxy(0).X = 14:   Zxy(0).Y = 28
Zxy(1).X = 26:   Zxy(1).Y = 42
Zxy(2).X = 26:   Zxy(2).Y = 28

For i = 0 To 2
    Zxy(i).X = (Dev_PIC(LJL_IDX).Left + Dev_PIC(LJL_IDX).Width * Zxy(i).X / LJL_W) * Window_Scale_X
    Zxy(i).Y = (Dev_PIC(LJL_IDX).Top + Dev_PIC(LJL_IDX).Height * Zxy(i).Y / LJL_H) * Window_Scale_Y
Next i
Select Case LJL(index).JLM_OPEN_FM_Style
    Case 0      'Xhd
        V = 0
        If (Dev_XHD(LJL(index).JLM_OPEN_FM_IDX).Value = 1) And (Dev_XHD(LJL(index).JLM_CLOSE_FM_IDX).Value = 0) Then
            V = 1       'Open
        Else
            If (Dev_XHD(LJL(index).JLM_OPEN_FM_IDX).Value = 0) And (Dev_XHD(LJL(index).JLM_CLOSE_FM_IDX).Value = 1) Then
                V = 0   'Close
            End If
        End If
        LJL(index).JLM_Value = V
    Case 1      'Pic
        V = 0
        If (Dev_PIC(LJL(index).JLM_OPEN_FM_IDX).Value = 1) And (Dev_PIC(LJL(index).JLM_CLOSE_FM_IDX).Value = 0) Then
            V = 1       'High
        Else
            If (Dev_PIC(LJL(index).JLM_OPEN_FM_IDX).Value = 0) And (Dev_PIC(LJL(index).JLM_CLOSE_FM_IDX).Value = 1) Then
                V = 0   'Low
            End If
        End If
        LJL(index).JLM_Value = V
End Select

If LJL(index).Last_JLM_Status <> LJL(index).JLM_Value Then
    Val_Chg = True
Else
    Val_Chg = False
End If
If Val_Chg Or Force_Show Then
    If LJL(index).JLM_Value = 1 Then    'high
        F_ShowLc.ForeColor = QBColor(0)
        F_ShowLc.FillStyle = 0
        F_ShowLc.FillColor = QBColor(14)    'Light Yellow
        SetPolyFillMode F_ShowLc.hdc, 2
        Polygon F_ShowLc.hdc, Zxy(0), 3
    Else
        If LJL(index).JLM_Value = 0 Then 'Low
            F_ShowLc.ForeColor = QBColor(0)
            F_ShowLc.FillStyle = 0
            F_ShowLc.FillColor = F_ShowLc.BackColor
            SetPolyFillMode F_ShowLc.hdc, 2
            Polygon F_ShowLc.hdc, Zxy(0), 3
        End If
    End If
    LJL(index).Last_JLM_Status = LJL(index).JLM_Value
End If
Exit Function
Run_ljl_error:
           Call write_error(3001, "运行指定密炼机的加料门错误")
End Function

'***************************************************************************
'*  函数功能：由设备的名称（NAME）及设备类型（STYLE）获取设备的索引号         *
'***************************************************************************
Public Function Get_Index_From_Name(Name As String, Style As Integer) As Integer
Dim i As Integer
Get_Index_From_Name = -1
If Style = 1 Then   'FaMen
    For i = 1 To Pic_num
        If Trim(Dev_PIC(i).Name) = Trim(Name) Then
            Get_Index_From_Name = i
            Exit For
        End If
    Next i
Else            'XHD
    For i = 1 To Xhd_num
        If Trim(Dev_XHD(i).Name) = Trim(Name) Then
            Get_Index_From_Name = i
            Exit For
        End If
    Next i
End If
End Function

'****************************************************************************
'*  函数功能：运行指定（INDEX）的密炼机（FORCE＝TRUE 强行显示数据）（定时器用）*
'****************************************************************************
Public Function Run_LJL(index As Integer, Force_Show As Boolean)
Dim IDX As Integer
Dim Val_Chg As Boolean
IDX = -1
For i = 0 To LJL_Num - 1
    If LJL(i).index = index Then
        IDX = i:        Exit For
    End If
Next i

If IDX <> -1 Then
    Call Run_LJL_SDS(index, IDX, Force_Show)    '显示上顶栓
    Call Run_LJL_JLM(index, IDX, Force_Show)    '显示加料门
    Call Run_LJL_XDS(index, IDX, Force_Show)    '显示下顶栓
End If
End Function

'****************************************************************************
'*  函数功能：演示指定图库名（TK_NAME）的生产过程                             *
'*          change by zgh               tk_name  change mathine
'****************************************************************************
Public Function ReadDb_Load_Tk(Tk_Name As String)
'Public Function ReadDb_Load_Tk(ByVal TK_Mathine As Integer)
On Error Resume Next
Dim db As Database, rs As Recordset
Dim Name As String
Dim answer As Integer
'Dim Tk_Name  As String

Set db = Workspaces(0).OpenDatabase(Runtk_DataBase_Path + "\" + "allmdb.mdb")
'Set rs = db.OpenRecordset("select  tk_name from tk_machine where  machine=" & TK_Mathine, dbOpenDynaset)
        'rs.FindFirst "mathine=" & tk_mathine
 '       If rs.EOF Then
 '                   Call Speak_Error("不能打开图库")
  '                  rs.Close
    '                db.Close
   '                 Exit Function
     '   End If
      '  Tk_Name = rs!Tk_Name
'rs.Close

Set rs = db.OpenRecordset("auth_table", dbOpenDynaset)
rs.FindFirst "trim(name)='" + Trim(Tk_Name) + "'"
If Not rs.NoMatch Then
      If F_ShowLc.Visible = True Then
          Unload F_ShowLc
          Load F_ShowLc
      End If
      If IsNull(rs!Mathine) Then
         TK_Mathine = 0
      Else
         TK_Mathine = rs!Mathine
      End If
      
      F_ShowLc.BackColor = rs!back_color
      F_ShowLc.Left = 0
      F_ShowLc.Top = 0
      F_ShowLc.Width = rs!Width
      F_ShowLc.Height = rs!Height
      Old_Window_Width = F_ShowLc.Width
      Old_Window_Height = F_ShowLc.Height
      
      'F_ShowLc.ScaleWidth = PixelX(F_ShowLc.Width)
      'F_ShowLc.ScaleHeight = PixelX(F_ShowLc.Height)
      Old_Window_ScaleWidth = F_ShowLc.ScaleWidth
      Old_Window_ScaleHeight = F_ShowLc.ScaleHeight
      
        If (rs!statuspnl_left <> -1) And (rs!statuspnl_top <> -1) Then
            StatusPanel.Left = PixelX(rs!statuspnl_left)
            StatusPanel.Top = PixelY(rs!statuspnl_top)
            StatusPanel.Visible = True
        Else
            StatusPanel.Visible = False
        End If
              
      'KJ(0).BackColor = F_ShowLc.BackColor
      Gqg.Color = rs!q_color: Gqg.Width = rs!q_width
      Gyg.Color = rs!y_color: Gyg.Width = rs!y_width
      Glg.Color = rs!j_color: Glg.Width = rs!j_width
      
      Cur_FileName = Trim(rs!Name)
      If InStr(1, MAIN_MDI.Caption, "-- 正在演示流程图:", 1) = 0 Then
            If F_ShowLc.Tag <> "" Then
                MAIN_MDI.Caption = MAIN_MDI.Caption + " -- 正在演示流程图: " + Trim(Cur_FileName) + " (" + F_ShowLc.Tag + "号机组)"
            Else
                MAIN_MDI.Caption = MAIN_MDI.Caption + " -- 正在演示流程图: " + Trim(Cur_FileName)
            End If
      Else
            Tmp = InStr(1, MAIN_MDI.Caption, "-- 正在演示流程图:", 1)
            If F_ShowLc.Tag <> "" Then
                MAIN_MDI.Caption = Left(MAIN_MDI.Caption, Tmp - 1) + "-- 正在演示流程图: " + Cur_FileName + " (" + F_ShowLc.Tag + "号机组)"
            Else
                MAIN_MDI.Caption = Left(MAIN_MDI.Caption, Tmp - 1) + "-- 正在演示流程图: " + Cur_FileName
            End If
      End If
      F_ShowLc.Show
      Call ReadDb_Load_From_Mdb(Trim(rs!pathname))
      F_ShowLc.SetFocus
Else
      Call Speak_Error("未找到图库 :" + "'" + Trim(Tk_Name) + "'")
End If

 rs.Close
db.Close

End Function

'****************************************************************************
'*  函数功能：显示指定设备（INDEX）的外观                                     *
'****************************************************************************
Public Function ReadDb_Draw_Pic(index As Integer)
Dim i, j As Long
Dim L, T, H, W As Long
Static Pre_Pic_Style, Pre_Pic_Pos As Integer

If (Dev_PIC(index).Action = 1) Then Exit Function     'action=1 HIDE
If (Dev_PIC(index).Style <> Pre_Pic_Style) Or (Dev_PIC(index).Position <> Pre_Pic_Pos) Then
    Select Case Dev_PIC(index).Position
        Case 0          'left
            PicDemo.DataField = "ComPicLeft"
            PicMask.DataField = "MaskPicLeft"
        Case 1          'up
            PicDemo.DataField = "ComPicUp"
            PicMask.DataField = "MaskPicUp"
        Case 2          'right
            PicDemo.DataField = "ComPicRight"
            PicMask.DataField = "MaskPicRight"
        Case 3          'down
            PicDemo.DataField = "ComPicDown"
            PicMask.DataField = "MaskPicDown"
    End Select
    
    'Make PicDemo's White color to Black color
    For i = 0 To Me.PicDemo.ScaleWidth - 1
        For j = 0 To Me.PicDemo.ScaleHeight - 1
            If Me.PicDemo.Point(i, j) = QBColor(15) Then 'Point is White
                Me.PicDemo.FillColor = QBColor(0)
                Me.PicDemo.FillStyle = 0
                Call ExtFloodFill(Me.PicDemo.hdc, i, j, QBColor(15), FLOODFILLSURFACE)
            End If
        Next j
    Next i
    Pre_Pic_Style = Dev_PIC(index).Style
    Pre_Pic_Pos = Dev_PIC(index).Position
End If
    
    L = Dev_PIC(index).Left:    T = Dev_PIC(index).Top
    W = Dev_PIC(index).Width:    H = Dev_PIC(index).Height
    If Me.PicMask.Picture <> 0 Then
        Me.PaintPicture Me.PicMask.Image, L, T, W, H, 0, 0, Me.PicMask.Width, Me.PicMask.Height, vbSrcAnd
        Me.PaintPicture Me.PicDemo.Image, L, T, W, H, 0, 0, Me.PicDemo.Width, Me.PicDemo.Height, vbSrcInvert
    End If
End Function

'****************************************************************************
'*  函数功能：从数据库（图库名所对应的数据库名MDB_NAME）取出设备数据           *
'****************************************************************************
Public Function ReadDb_Load_From_Mdb(Mdb_name As String)
Dim Loc_DB, Comm_Db As Database
Dim Loc_RS, RL_RS, SL_RS, Comm_RS As Recordset
Dim i, j, Zxy_index As Integer
Dim head, sqlstr As String
If (Trim(Mdb_name) = "") Or (Dir$(Trim(Mdb_name)) = Null) Then
  Exit Function
End If
F_ShowLc.Timer1 = False
F_ShowLc.Timer2 = False

F_ShowLc.MousePointer = 11
Set Loc_DB = Workspaces(0).OpenDatabase(Trim(Mdb_name))
Set Loc_RS = Loc_DB.OpenRecordset("grap_table", dbOpenDynaset)
Set RL_RS = Loc_DB.OpenRecordset("relation_table", dbOpenDynaset)
Set SL_RS = Loc_DB.OpenRecordset("sub_line_table", dbOpenDynaset)

'读取图象数据
Kj_num = 0                  '块胶控件数
Tl_num = 0
Pic_num = 0

Set Loc_RS = Loc_DB.OpenRecordset("Select * From Grap_Table Where Code Like 'O*'")

MAIN_MDI.Process_Msg.Visible = True
MAIN_MDI.ProcessBar.Visible = True
MAIN_MDI.Process_Msg.Caption = "装入图象:"
MAIN_MDI.ProcessBar.Min = 0
Loc_RS.MoveLast
MAIN_MDI.ProcessBar.Max = Loc_RS.RecordCount
MAIN_MDI.ProcessBar.Value = 0
DoEvents

Loc_RS.MoveFirst
Do Until Loc_RS.EOF
    Pic_num = Pic_num + 1
    Dev_PIC(Pic_num).Code = Trim(Loc_RS!Code)
    Dev_PIC(Pic_num).Name = Trim(Loc_RS!Name)
    Dev_PIC(Pic_num).Machine = Loc_RS!Machine
    Dev_PIC(Pic_num).Style = Loc_RS!Style
    Dev_PIC(Pic_num).Position = Loc_RS!Position
    Dev_PIC(Pic_num).Action = Loc_RS!Action
    
    Dev_PIC(Pic_num).Left = PixelX(Loc_RS!Left)
    Dev_PIC(Pic_num).Top = PixelY(Loc_RS!Top)
    Dev_PIC(Pic_num).Width = PixelX(Loc_RS!Width)
    Dev_PIC(Pic_num).Height = PixelY(Loc_RS!Height)
    Dev_PIC(Pic_num).Old_Left = PixelX(Loc_RS!Left)
    Dev_PIC(Pic_num).Old_Top = PixelY(Loc_RS!Top)
    Dev_PIC(Pic_num).Old_Width = PixelX(Loc_RS!Width)
    Dev_PIC(Pic_num).Old_Height = PixelY(Loc_RS!Height)
    Dev_PIC(Pic_num).Enabled = True
'    ReadDb_Draw_Pic (Pic_num)
'    If (Dev_PIC(Pic_num).Style = 6) Then  '胶料秤
'        Kj_num = Kj_num + 1
'        Load F_ShowLc.KJ(Kj_num)
'        F_ShowLc.KJ(Kj_num).Left = Dev_PIC(Pic_num).Left
'        F_ShowLc.KJ(Kj_num).Width = Dev_PIC(Pic_num).Width
'        F_ShowLc.KJ(Kj_num).Top = Dev_PIC(Pic_num).Top - F_ShowLc.KJ(Kj_num).Height
'        F_ShowLc.KJ(Kj_num).Tag = str(Pic_num)
'        F_ShowLc.KJ(Kj_num).Visible = True
'    End If
    
    O_num = O_num + 1
    Loc_RS.MoveNext
    
    MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
Loop

'===============================================================
'读取折线数据(relation_table,sub_line_table)
Zline_num = 0: Line_num = 0
J_num = 0: Y_num = 0: Q_num = 0
'供胶管（料）
    '由relation_table查找管道
    Set Loc_RS = Loc_DB.OpenRecordset("Select * From Grap_Table Where Code Like 'J*'")
    
    MAIN_MDI.Process_Msg.Caption = "装入供胶管:"
    MAIN_MDI.ProcessBar.Min = 0
    Loc_RS.MoveLast
    MAIN_MDI.ProcessBar.Max = Loc_RS.RecordCount
    MAIN_MDI.ProcessBar.Value = 0
    DoEvents
    
    On Error Resume Next
    Loc_RS.MoveFirst
    Do Until Loc_RS.EOF
        Dev_ZLine(Zline_num).Enabled = True
        Dev_ZLine(Zline_num).Code = Trim(Loc_RS!Code)
        Dev_ZLine(Zline_num).Name = Trim(Loc_RS!Name)
        Dev_ZLine(Zline_num).Style = Loc_RS!Style
                
        'start_pic_index,   end_pic_index
        RL_RS.FindFirst "Code='" + Trim(Dev_ZLine(Zline_num).Code) + "'"
        If Not RL_RS.NoMatch Then
            For i = 1 To Pic_num
                If Trim(Dev_PIC(i).Code) = Trim(RL_RS!last_code) Then
                    Dev_ZLine(Zline_num).Start_Pic_Index = i
                    Exit For
                End If
            Next i
            For i = 1 To Pic_num
                If Trim(Dev_PIC(i).Code) = Trim(RL_RS!next_code) Then
                    Dev_ZLine(Zline_num).End_Pic_Index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        Zxy_index = 0
        sqlstr = "Select * From sub_line_Table Where Code ='" + Trim(Dev_ZLine(Zline_num).Code) + "' order by index"
        Set SL_RS = Loc_DB.OpenRecordset(sqlstr)
        SL_RS.MoveFirst
        Do Until SL_RS.EOF
 '           F_ShowLc.DrawWidth = Glg.Width
 '           F_ShowLc.Line -(PixelX(Sl_RS!Left), PixelY(Sl_RS!Top)), Glg.Color
            Dev_ZLine(Zline_num).Zxy(Zxy_index, 0) = PixelX(SL_RS!Left)
            Dev_ZLine(Zline_num).Zxy(Zxy_index, 1) = PixelY(SL_RS!Top)
            Zxy_index = Zxy_index + 1
            SL_RS.MoveNext
        Loop
        Dev_ZLine(Zline_num).Zdot_num = Zxy_index
        
        J_num = J_num + 1
        Zline_num = Zline_num + 1
        Loc_RS.MoveNext
        
        MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
    Loop
  
'供汽管
    '由relation_table查找管道
    Set Loc_RS = Loc_DB.OpenRecordset("Select * From Grap_Table Where Code Like 'Q*'")
    
    MAIN_MDI.Process_Msg.Caption = "装入供汽管:"
    MAIN_MDI.ProcessBar.Min = 0
    Loc_RS.MoveLast
    MAIN_MDI.ProcessBar.Max = Loc_RS.RecordCount
    MAIN_MDI.ProcessBar.Value = 0
    DoEvents

    On Error Resume Next
    Loc_RS.MoveFirst
    Do Until Loc_RS.EOF
        Dev_ZLine(Zline_num).Code = Trim(Loc_RS!Code)
        Dev_ZLine(Zline_num).Name = Trim(Loc_RS!Name)
        Dev_ZLine(Zline_num).Style = Loc_RS!Style
        Dev_ZLine(Zline_num).Enabled = True
        
        'start_pic_index,   end_pic_index
        RL_RS.FindFirst "code='" + Trim(Dev_ZLine(Zline_num).Code) + "'"
        If Not RL_RS.NoMatch Then
            For i = 1 To Pic_num
                If Trim(Dev_PIC(i).Code) = Trim(RL_RS!last_code) Then
                    Dev_ZLine(Zline_num).Start_Pic_Index = i
                    Exit For
                End If
            Next i
            For i = 1 To Pic_num
                If Trim(Dev_PIC(i).Code) = Trim(RL_RS!next_code) Then
                    Dev_ZLine(Zline_num).End_Pic_Index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        Zxy_index = 0
        sqlstr = "Select * From sub_line_Table Where Code ='" + Trim(Dev_ZLine(Zline_num).Code) + "' order by index"
        Set SL_RS = Loc_DB.OpenRecordset(sqlstr)
        SL_RS.MoveFirst
        Do Until SL_RS.EOF
 '           F_ShowLc.DrawWidth = Glg.Width
 '           F_ShowLc.Line -(PixelX(Sl_RS!Left), PixelY(Sl_RS!Top)), Gqg.Color
            Dev_ZLine(Zline_num).Zxy(Zxy_index, 0) = PixelX(SL_RS!Left)
            Dev_ZLine(Zline_num).Zxy(Zxy_index, 1) = PixelY(SL_RS!Top)
            Zxy_index = Zxy_index + 1
            SL_RS.MoveNext
        Loop
        Dev_ZLine(Zline_num).Zdot_num = Zxy_index
        
        Q_num = Q_num + 1
        Zline_num = Zline_num + 1
        Loc_RS.MoveNext
    
        MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
    Loop
  
'供油管
    '由relation_table查找管道
    Set Loc_RS = Loc_DB.OpenRecordset("Select * From Grap_Table Where Code Like 'Y*'")
    
    MAIN_MDI.Process_Msg.Caption = "装入供油管:"
    MAIN_MDI.ProcessBar.Min = 0
    Loc_RS.MoveLast
    MAIN_MDI.ProcessBar.Max = Loc_RS.RecordCount
    MAIN_MDI.ProcessBar.Value = 0
    DoEvents

    On Error Resume Next
    Loc_RS.MoveFirst
    Do Until Loc_RS.EOF
        Dev_ZLine(Zline_num).Code = Trim(Loc_RS!Code)
        Dev_ZLine(Zline_num).Name = Trim(Loc_RS!Name)
        Dev_ZLine(Zline_num).Style = Loc_RS!Style
        Dev_ZLine(Zline_num).Enabled = True
        
        'start_pic_index,   end_pic_index
        RL_RS.FindFirst "code='" + Trim(Dev_ZLine(Zline_num).Code) + "'"
        If Not RL_RS.NoMatch Then
            For i = 1 To Pic_num
                If Trim(Dev_PIC(i).Code) = Trim(RL_RS!last_code) Then
                    Dev_ZLine(Zline_num).Start_Pic_Index = i
                    Exit For
                End If
            Next i
            For i = 1 To Pic_num
                If Trim(Dev_PIC(i).Code) = Trim(RL_RS!next_code) Then
                    Dev_ZLine(Zline_num).End_Pic_Index = i
                    Exit For
                End If
            Next i
        End If
        
        '由sub_line_table画线
        Zxy_index = 0
        sqlstr = "Select * From sub_line_Table Where Code ='" + Trim(Dev_ZLine(Zline_num).Code) + "' order by index"
        Set SL_RS = Loc_DB.OpenRecordset(sqlstr)
        SL_RS.MoveFirst
        Do Until SL_RS.EOF
 '           F_ShowLc.DrawWidth = Glg.Width
 '           F_ShowLc.Line -(PixelX(Sl_RS!Left), PixelY(Sl_RS!Top)), Gyg.Color
            Dev_ZLine(Zline_num).Zxy(Zxy_index, 0) = PixelX(SL_RS!Left)
            Dev_ZLine(Zline_num).Zxy(Zxy_index, 1) = PixelY(SL_RS!Top)
            Zxy_index = Zxy_index + 1
            SL_RS.MoveNext
        Loop
        Dev_ZLine(Zline_num).Zdot_num = Zxy_index
        
        Y_num = Y_num + 1
        Zline_num = Zline_num + 1
        Loc_RS.MoveNext
        
        MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
    Loop
    SL_RS.Close
    RL_RS.Close
    
'===============================================================
'读取信号灯数据(grap_table)
F_ShowLc.FillColor = QBColor(12)
F_ShowLc.ForeColor = QBColor(0)
F_ShowLc.DrawWidth = 1
Xhd_num = 0

Set Loc_RS = Loc_DB.OpenRecordset("Select * From Grap_Table Where Code Like 'X*'")

MAIN_MDI.Process_Msg.Caption = "装入信号灯:"
MAIN_MDI.ProcessBar.Min = 0
Loc_RS.MoveLast
MAIN_MDI.ProcessBar.Max = Loc_RS.RecordCount
MAIN_MDI.ProcessBar.Value = 0
DoEvents

Loc_RS.MoveFirst
Do Until Loc_RS.EOF
    Dev_XHD(Xhd_num).Enabled = True
    Xhd_num = Xhd_num + 1
    Dev_XHD(Xhd_num).Code = Trim(Loc_RS!Code)
    Dev_XHD(Xhd_num).Name = Trim(Loc_RS!Name)
    Dev_XHD(Xhd_num).Machine = Loc_RS!Machine
    Dev_XHD(Xhd_num).Action = Loc_RS!Action
    
    Dev_XHD(Xhd_num).Left = PixelX(Loc_RS!Left)
    Dev_XHD(Xhd_num).Top = PixelY(Loc_RS!Top)
    Dev_XHD(Xhd_num).Width = PixelX(Loc_RS!Width)
    Dev_XHD(Xhd_num).Height = PixelY(Loc_RS!Height)
    Dev_XHD(Xhd_num).Old_Left = PixelX(Loc_RS!Left)
    Dev_XHD(Xhd_num).Old_Top = PixelY(Loc_RS!Top)
    Dev_XHD(Xhd_num).Old_Width = PixelX(Loc_RS!Width)
    Dev_XHD(Xhd_num).Old_Height = PixelY(Loc_RS!Height)
'    Ellipse F_ShowLc.hdc, Dev_xhd(Xhd_num).Left, Dev_xhd(Xhd_num).Top, Dev_xhd(Xhd_num).Left + Dev_xhd(Xhd_num).Width, Dev_xhd(Xhd_num).Top + Dev_xhd(Xhd_num).Height

    X_num = X_num + 1
    Loc_RS.MoveNext
    
    MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
Loop

'---------------------------------------
'读取密炼机数据
LJL_Num = 0
Set Loc_RS = Loc_DB.OpenRecordset("Select * From Grap_Table Where style=9") 'LJL

MAIN_MDI.Process_Msg.Caption = "装入密炼机:"
MAIN_MDI.ProcessBar.Min = 0
Loc_RS.MoveLast
MAIN_MDI.ProcessBar.Max = Loc_RS.RecordCount
MAIN_MDI.ProcessBar.Value = 0
DoEvents

Loc_RS.MoveFirst
Do Until Loc_RS.EOF
    For j = 1 To Pic_num
        If (Trim(Dev_PIC(j).Name) = Trim(Loc_RS!Name)) And (Dev_PIC(j).Style = 9) Then
            LJL(LJL_Num).index = j
            LJL(LJL_Num).Name = Trim(Loc_RS!Name)
            Exit For
        End If
    Next j
    
    LJL(LJL_Num).SDS_High_Style = Val(Mid(Trim(Loc_RS!remark), 5, 1))
    LJL(LJL_Num).SDS_Mid_Style = Val(Mid(Trim(Loc_RS!remark), 11, 1))
    LJL(LJL_Num).SDS_Low_Style = Val(Mid(Trim(Loc_RS!remark), 17, 1))
    LJL(LJL_Num).JLM_OPEN_FM_Style = Val(Mid(Trim(Loc_RS!remark), 23, 1))
    LJL(LJL_Num).JLM_CLOSE_FM_Style = Val(Mid(Trim(Loc_RS!remark), 29, 1))
    LJL(LJL_Num).XDS_OPEN_FM_Style = Val(Mid(Trim(Loc_RS!remark), 35, 1))
    LJL(LJL_Num).XDS_CLOSE_FM_Style = Val(Mid(Trim(Loc_RS!remark), 41, 1))
    
    LJL(LJL_Num).SDS_High_IDX = Get_Index_From_Name(Mid(Trim(Loc_RS!remark), 1, 4), LJL(LJL_Num).SDS_High_Style)
    LJL(LJL_Num).SDS_Mid_IDX = Get_Index_From_Name(Mid(Trim(Loc_RS!remark), 7, 4), LJL(LJL_Num).SDS_Mid_Style)
    LJL(LJL_Num).SDS_Low_IDX = Get_Index_From_Name(Mid(Trim(Loc_RS!remark), 13, 4), LJL(LJL_Num).SDS_Low_Style)
    LJL(LJL_Num).JLM_OPEN_FM_IDX = Get_Index_From_Name(Mid(Trim(Loc_RS!remark), 19, 4), LJL(LJL_Num).JLM_OPEN_FM_Style)
    LJL(LJL_Num).JLM_CLOSE_FM_IDX = Get_Index_From_Name(Mid(Trim(Loc_RS!remark), 25, 4), LJL(LJL_Num).JLM_CLOSE_FM_Style)
    LJL(LJL_Num).XDS_OPEN_FM_IDX = Get_Index_From_Name(Mid(Trim(Loc_RS!remark), 31, 4), LJL(LJL_Num).XDS_OPEN_FM_Style)
    LJL(LJL_Num).XDS_CLOSE_FM_IDX = Get_Index_From_Name(Mid(Trim(Loc_RS!remark), 37, 4), LJL(LJL_Num).XDS_CLOSE_FM_Style)
    
    LJL_Num = LJL_Num + 1
    Loc_RS.MoveNext     'LJL
    
    MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
Loop
'---------------------------------------
Loc_RS.Close
    
'读取文字数据 (prompt_table)
    i = 0
    Txt_num = 0
    Status_txt_num = 0
    Set Loc_RS = Loc_DB.OpenRecordset("prompt_table", dbOpenDynaset)
    
    MAIN_MDI.Process_Msg.Caption = "装入文字:"
    MAIN_MDI.ProcessBar.Min = 0
    Loc_RS.MoveLast
    MAIN_MDI.ProcessBar.Max = Loc_RS.RecordCount
    MAIN_MDI.ProcessBar.Value = 0
    DoEvents
    
    On Error Resume Next
    Loc_RS.MoveFirst
    Do Until Loc_RS.EOF
      If Loc_RS!Style = 0 Then          'Prompt Text
        Txt_num = Txt_num + 1
        Load F_ShowLc!txt(Txt_num)
        F_ShowLc!txt(Txt_num).Caption = Trim(Loc_RS!Prompt)
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
      Else              'Status Text
        Status_txt_num = Status_txt_num + 1
        Load F_ShowLc!stxt(Status_txt_num)
        F_ShowLc!stxt(Status_txt_num).Caption = Trim(Loc_RS!Prompt)
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
        'F_ShowLc!stxt(Status_txt_num).Visible = True
        'F_ShowLc!stxt(Status_txt_num).ZOrder 0
        
        Status_Txt(Status_txt_num).Relation_code = Trim(Loc_RS!Relation_code)
      End If
      Loc_RS.MoveNext
      
      MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
  Loop
  Loc_RS.Close
  Loc_DB.Close

'设置碳黑秤（THC）,油料秤（YLC）的min,max_size值
Set Comm_Db = Workspaces(0).OpenDatabase("c:\ljxt\comm.mdb")
Set Comm_RS = Comm_Db.OpenRecordset("send_table", dbOpenDynaset)
For k = 1 To Pic_num
    If Dev_PIC(k).Style = 7 Then
        Comm_RS.FindFirst "mathine=" + CStr(Dev_PIC(k).Machine) + " and code=" + Left(Trim(Dev_PIC(k).Name), 1)
         If Not Comm_RS.NoMatch Then
             Dev_PIC(k).Min = Comm_RS!Min
             Dev_PIC(k).Max_Size = Comm_RS!Max
         End If
    End If
Next k
Comm_RS.Close
Comm_Db.Close
  
 ' For i = 1 To Txt_num
 '   F_ShowLc.ForeColor = txt(i).ForeColor
 '   F_ShowLc.Font = txt(i).Font
 '   F_ShowLc.FontName = txt(i).FontName
 '   F_ShowLc.FontSize = txt(i).FontSize
 '   Me.CurrentX = txt(i).Left
 '   Me.CurrentY = txt(i).Top
    
 '   F_ShowLc.Print txt(i).Caption

''    F_ShowLc!txt(i).Visible = True
''    F_ShowLc!txt(i).ZOrder 0
'  Next i

'  For i = 1 To Status_txt_num
'    F_ShowLc!stxt(i).Visible = True
'    F_ShowLc!stxt(i).ZOrder 0
'  Next i

'-------------------------------
For i = 0 To Zline_num - 1
    If Left(Dev_ZLine(i).Code, 1) <> "Q" Then
        Dev_ZLine(i).Can_Comein = True
        Dev_ZLine(i).Can_Goout = True
    End If
Next i
Call Form_Paint

MAIN_MDI.Process_Msg.Visible = False
MAIN_MDI.ProcessBar.Visible = False
F_ShowLc.Timer1 = True
F_ShowLc.Timer2 = True

F_ShowLc.MousePointer = 0
End Function

'****************************************************************************
'*  函数功能：动态仿真用的流体网络算法（算法：图的深度优先递归）（定时器用）    *
'****************************************************************************
Public Function Luo_DFS(v0 As Integer)
Dim i As Integer
Dim W As Integer
    '记录折线路径
    If Arcs(Last_Vtx, v0) <> -1 Then
        Road.ZIndex(Road.Length) = Arcs(Last_Vtx, v0)
        Road.Length = Road.Length + 1
    End If
    Visited(v0) = True
    Passed_vtx(v0) = True
    '------------------------
    Last_Vtx = v0
    W = Luo_FirstAdj(v0)            '获取第一个邻接点
    While W <> -1                   '当邻接点存在时
       If Not Visited(W) Then
            Search_new_road = True
            Last_Vtx = v0
            Luo_DFS (W)             '递归
        Else        '存在环路
            If Arcs(v0, W) <> -1 Then
                Road.ZIndex(Road.Length) = Arcs(v0, W)
                Road.Length = Road.Length + 1
            End If
            Luo_Set_Road            '设置当前路径中的设备参数
        
            For k = 0 To Road.Length - 1
                If Dev_ZLine(Road.ZIndex(k)).End_Pic_Index = v0 Then
                    Road.Length = k + 1
                    Exit For
                End If
            Next k
            Visited(v0) = False
            Exit Function
        End If
        Last_Vtx = v0
        W = Luo_NextAdj(v0, W)  '找下一邻接点
    Wend
    '回搠
    If Search_new_road Then
        Luo_Set_Road    '找完一条路径,分析、设置路径中的设备参数
        Search_new_road = False
    End If
    Visited(v0) = False
    If Road.Length > 0 Then
        Road.Length = Road.Length - 1
    End If
End Function

'****************************************************************************
'*  函数功能：查找V接点的第一个邻接点                                         *
'****************************************************************************
Public Function Luo_FirstAdj(V As Integer) As Integer
Dim i As Integer
    Luo_FirstAdj = -1
    For i = 1 To Pic_num
        If Arcs(V, i) <> -1 Then    '当邻接点存在时
            Luo_FirstAdj = i
            Exit Function
        End If
    Next i
End Function

'****************************************************************************
'*  函数功能：判断V接点是否为起点                                             *
'****************************************************************************
Public Function DFS_Is_Start_Vtx(V As Integer) As Boolean
Dim i As Integer
DFS_Is_Start_Vtx = True
For i = 1 To Pic_num
    If Arcs(i, V) <> -1 Then
        DFS_Is_Start_Vtx = False
        Exit For
    End If
Next i
End Function


'****************************************************************************
'*  函数功能：查找V接点的下一个邻接点                                         *
'****************************************************************************
Public Function Luo_NextAdj(V As Integer, W As Integer) As Integer
Dim i As Integer
    Luo_NextAdj = -1
    For i = W + 1 To Pic_num
        If Arcs(V, i) <> -1 Then    '当下一个邻接点存在时
            Luo_NextAdj = i
            Exit Function
        End If
    Next i
End Function

'****************************************************************************
'*  函数功能：对整个图库执行流体网络算法，设置所有设备的运行参数（定时器用）    *
'****************************************************************************
Public Function Luo_Traver()
Dim vi  As Integer
Dim Exit_Do As Boolean
Dim Visited_Vtx(1 To 100) As Integer

For vi = 1 To Pic_num
    Visited_Vtx(vi) = False:    Visited(vi) = False:    Passed_vtx(vi) = False
Next vi
For vi = 0 To Zline_num - 1
    Zline_obj_passing(vi) = False
Next vi
Do
    Exit_Do = True
    For vi = 1 To Pic_num
'        If (Not visited(vi)) And (dfs_is_start_vtx(vi) And (luo_firstadj(vi) <> -1)) Then
         If (Not Visited_Vtx(vi)) And (Not Passed_vtx(vi)) And (DFS_Is_Start_Vtx(vi) And (Luo_FirstAdj(vi) <> -1)) Then
            Visited_Vtx(vi) = True
            Last_Vtx = vi
            Road.Length = 0
            Luo_DFS (vi)            '执行流体网络算法
            Exit_Do = False
            Exit For
        End If
    Next vi
Loop Until Exit_Do

'环路遍历
Do
    Exit_Do = True
    For vi = 1 To Pic_num
         If (Not Visited_Vtx(vi)) And (Not Passed_vtx(vi)) And (Luo_FirstAdj(vi) <> -1) Then
            Passed_vtx(vi) = True
            Last_Vtx = vi
            Road.Length = 0
            Luo_DFS (vi)            '执行流体网络算法
            Exit_Do = False
            Exit For
        End If
    Next vi
Loop Until Exit_Do

For vi = 0 To Zline_num - 1         '设置管道参数
    If Dev_ZLine(vi).Obj_Passing <> Zline_obj_passing(vi) Then
        Dev_ZLine(vi).Obj_Passing = Zline_obj_passing(vi)
    End If
Next vi
End Function

'****************************************************************************
'*  函数功能：主调函数（设置设备初始参数）                                    *
'****************************************************************************
Sub Main(ByVal Do_mathine As Integer)
Runtk_DataBase_Path = "c:\ljxt\tkbmp"
Window_Scale_X = 1:      Window_Scale_Y = 1
Dls.Space_len = 6
Dls.Xd_len = 24
Glg.Neijing_width = 1
Glg.Obj_color = QBColor(15)
Gqg.Neijing_width = 1
Gqg.Obj_color = QBColor(14)
Gyg.Neijing_width = 1
Gyg.Obj_color = QBColor(14)

'========阀门做动画用==========
'Dir Left：向上
Fm_hole(0).scale_x1 = 1# / 21#:     Fm_hole(0).scale_y1 = 1# / 23#
Fm_hole(0).scale_w1 = 8# / 21#:     Fm_hole(0).scale_h1 = 8# / 23#
Fm_hole(0).scale_x2 = 1# / 21#:     Fm_hole(0).scale_y2 = 13# / 23#
Fm_hole(0).scale_w2 = 8# / 21#:     Fm_hole(0).scale_h2 = 8# / 23#
'Dir Up：向下
Fm_hole(1).scale_x1 = 1# / 23#:     Fm_hole(1).scale_y1 = 1# / 21#
Fm_hole(1).scale_w1 = 8# / 23#:     Fm_hole(1).scale_h1 = 8# / 21#
Fm_hole(1).scale_x2 = 13# / 23#:     Fm_hole(1).scale_y2 = 1# / 21#
Fm_hole(1).scale_w2 = 8# / 23#:     Fm_hole(1).scale_h2 = 8# / 21#
'Dir Right：向右
Fm_hole(2).scale_x1 = 11# / 21#:     Fm_hole(2).scale_y1 = 1# / 23#
Fm_hole(2).scale_w1 = 8# / 21#:     Fm_hole(2).scale_h1 = 8# / 23#
Fm_hole(2).scale_x2 = 11# / 21#:     Fm_hole(2).scale_y2 = 13# / 23#
Fm_hole(2).scale_w2 = 8# / 21#:     Fm_hole(2).scale_h2 = 8# / 23#
'Dir Down：向下
Fm_hole(3).scale_x1 = 1# / 23#:     Fm_hole(3).scale_y1 = 11# / 21#
Fm_hole(3).scale_w1 = 8# / 23#:     Fm_hole(3).scale_h1 = 8# / 21#
Fm_hole(3).scale_x2 = 13# / 23#:     Fm_hole(3).scale_y2 = 11# / 21#
Fm_hole(3).scale_w2 = 8# / 23#:     Fm_hole(3).scale_h2 = 8# / 21#

'调入演示系统图（GRAP_NAME）
Call ReadDb_Load_Tk(Grap_Name)
'Call ReadDb_Load_Tk(Do_mathine)
'将折线数据转化为流体网络用的邻接矩阵数据
Call DFS_Zline_To_Arcs

'使能设备
For i = 1 To Pic_num: Dev_PIC(i).Enabled = True: Next i
Set_Pic_Enabled 3, True
Set_Pic_Enabled 4, True
Set_Pic_Enabled 6, True
Set_Pic_Enabled 7, True
Set_Pic_Enabled 8, True
Set_Pic_Enabled 9, True
Set_Pic_Enabled 10, True
Set_Pic_Enabled 14, True
Set_Pic_Enabled 17, True
'set_pic_enabled 18, True
Set_Pic_Enabled 19, True
Road.Length = 0

For i = 1 To Pic_num
    Select Case Dev_PIC(i).Style
        Case 7  '碳黑秤、油料秤
            For j = 1 To Pic_num
                If (Dev_PIC(j).Style = 18) And (Dev_PIC(j).Machine = Dev_PIC(i).Machine) And (Left(Dev_PIC(j).Name, 1) = Left(Dev_PIC(i).Name, 1)) Then
                    Dev_PIC(j).Min = Dev_PIC(i).Min
                    Dev_PIC(j).Max_Size = Dev_PIC(i).Max_Size
                End If
            Next j
        Case 9  '密炼机
            Dev_PIC(i).Last_Status = -1
    End Select
    Dev_PIC(i).Beizhu = ""
Next i
For i = 0 To Zline_num - 1
    Dev_ZLine(i).Cur_Offset = 0
    Dev_ZLine(i).Speed = LS_NORMAL
Next i

'错误显示窗口初始化
With ErrorGrid
    .Height = .RowHeight(0) * 2 + Screen.TwipsPerPixelY * 2
    .Rows = 2
    .Cols = 5
    .ColWidth(0) = 35 * Screen.TwipsPerPixelX
    .ColWidth(1) = 45 * Screen.TwipsPerPixelX
    .ColWidth(2) = 180 * Screen.TwipsPerPixelX
    .ColWidth(3) = 45 * Screen.TwipsPerPixelX
    .ColWidth(4) = 65 * Screen.TwipsPerPixelX
    .FixedRows = 1
    .row = 0
    .col = 0:       .Text = "序号"
    .col = 1:       .Text = "编  号"
    .col = 2:       .Text = "错误提示"
    .col = 3:       .Text = "时  间"
    .col = 4:       .Text = "日  期"
End With

Last_Errors.Length = 0
Cur_Errors.Length = 0
DoEvents
End Sub

'****************************************************************************
'*  函数功能：运行指定（INDEX）的碳黑秤、油料秤（定时器用）                    *
'****************************************************************************
Public Function Run_THC_YLC(index As Integer, Force_Show As Boolean)
Dim i As Integer
Dim Bot_X, Bot_Y, High, K1, K2 As Double
Dim Wx, Wy, Ww, Wh, Fx, Fy, X1, Y1, X2, Y2, X3, Y3 As Long
Dim Ffc, Bc As Long
Dim Val_Chg As Boolean
Dim Scale_X, Scale_Y As Double
Dim Zxy(0 To 3) As POINTAPI
Dim Tmp_Val As Single
If Dev_PIC(index).Max_Size <= 0 Then
    Exit Function
End If
Fc = QBColor(5): Bc = QBColor(0)
Bot_X = 55# / 110#:     Bot_Y = 81# / 83#:     High = (81# - 27#) / 83#
'k1 = (16# - 44#) / (13# - 30#): k2 = (16# - 44#) / (49# - 30#)

Wx = Dev_PIC(index).Left '/ Window_Scale_X
Wy = Dev_PIC(index).Top '/ Window_Scale_Y
Ww = Dev_PIC(index).Width '/ Window_Scale_X
Wh = Dev_PIC(index).Height '/ Window_Scale_Y

Scale_X = Dev_PIC(index).Width / 110#
Scale_Y = Dev_PIC(index).Height / 83#
K1 = (1# - 55#) * Scale_Y / (27# - 81#) / Scale_X
K2 = (109# - 55#) * Scale_Y / (27# - 81#) / Scale_X

Val_Chg = IIf((Dev_PIC(index).Last_Status <> Dev_PIC(index).Value) Or (Dev_PIC(index).Last_Beizhu <> Dev_PIC(index).Beizhu), True, False)
If Dev_PIC(index).Value > Dev_PIC(index).Max_Size Then
    Tmp_Val = Dev_PIC(index).Max_Size
Else
    If Dev_PIC(index).Value < 0 Then
        Tmp_Val = 0
    Else
        Tmp_Val = Dev_PIC(index).Value
    End If
End If

If (Val_Chg Or Force_Show) Then
    DrawWidth = 1
    'draw corner
    X1 = Wx + Ww * Bot_X:       Y1 = Wy + Wh * Bot_Y
    Y2 = Y1 - Wh * High:            X2 = (Y2 - Y1) / K1 + X1
    X3 = (Y2 - Y1) / K2 + X1:       Y3 = Y2
    
    '顶边
    Zxy(0).X = CLng(X2 * Window_Scale_X):    Zxy(0).Y = CLng(Y2 * Window_Scale_Y)
    Zxy(3).X = CLng(X3 * Window_Scale_X):    Zxy(3).Y = CLng(Y2 * Window_Scale_Y)
    
    Tmp_Y = CDbl(Y1) - Wh * High * CDbl(Tmp_Val) / CDbl(Dev_PIC(index).Max_Size)
    'Tmp_Y = CDbl(Y1) - Wh * High * Exp(Log(CDbl(Tmp_Val) / CDbl(Dev_pic(index).Max_Size)) / 3)
    '底边
    Tmp_X = CDbl(Tmp_Y - Y1) / K1 + X1
    Zxy(1).X = CLng(Tmp_X * Window_Scale_X):    Zxy(1).Y = CLng(Tmp_Y * Window_Scale_Y)
    Tmp_X = CDbl(Tmp_Y - Y1) / K2 + X1
    Zxy(2).X = CLng(Tmp_X * Window_Scale_X):    Zxy(2).Y = CLng(Tmp_Y * Window_Scale_Y)

    '填充液面以上空间
    F_ShowLc.ForeColor = QBColor(0)
    F_ShowLc.FillStyle = 0
    F_ShowLc.FillColor = F_ShowLc.BackColor
    
    SetPolyFillMode F_ShowLc.hdc, 2
    Polygon F_ShowLc.hdc, Zxy(0), 4
    
    '填充液体
    '顶边
    Zxy(0).X = Zxy(1).X:     Zxy(0).Y = Zxy(1).Y
    Zxy(1).X = Zxy(2).X:     Zxy(1).Y = Zxy(2).Y
    '底部
    Zxy(2).X = CLng(X1 * Window_Scale_X):    Zxy(2).Y = CLng(Y1 * Window_Scale_Y)
    
    F_ShowLc.ForeColor = QBColor(0)
    F_ShowLc.FillStyle = 0
    F_ShowLc.FillColor = Fc
    
    SetPolyFillMode F_ShowLc.hdc, 2
    Polygon F_ShowLc.hdc, Zxy(0), 3

    '显示状态
    For j = 1 To Status_txt_num
        If Trim(Status_Txt(j).Relation_code) = Trim(Dev_PIC(index).Code) Then
            stxt(j).Caption = CStr(Dev_PIC(index).Value)
            
            With F_ShowLc
                FontName = "Fixedsys" 'stxt(j).FontName
                FontSize = stxt(j).FontSize
                ForeColor = stxt(j).ForeColor
                CurrentX = stxt(j).Left: CurrentY = stxt(j).Top
                'Print "----------"
                'CurrentX = stxt(j).Left: CurrentY = stxt(j).Top
                If Len(Trim(Dev_PIC(index).Beizhu)) > 0 Then
                    T = 11 - (Len(stxt(j).Caption) + 1 + 2 * Len(Trim(Dev_PIC(index).Beizhu)))
                    T = IIf(T >= 0, T, 0)
                    'stxt(j).Caption = String(Int(T / 2), "-") + stxt(j).Caption + "-" + Trim(Dev_pic(index).Beizhu) + String(T - Int(T / 2), "-")
                    ForeColor = Me.BackColor: Print String(Int(T / 2), "-");
                    ForeColor = stxt(j).ForeColor: Print stxt(j).Caption;
                    ForeColor = Me.BackColor: Print "-";
                    ForeColor = stxt(j).ForeColor: Print Trim(Dev_PIC(index).Beizhu);
                    ForeColor = Me.BackColor: Print String(T - Int(T / 2), "-");
                Else
                    T = 11 - Len(stxt(j).Caption)
                    T = IIf(T >= 0, T, 0)
                    'stxt(j).Caption = String(Int(T / 2), "-") + stxt(j).Caption + String(T - Int(T / 2), "-")
                    ForeColor = Me.BackColor: Print String(Int(T / 2), "-");
                    ForeColor = stxt(j).ForeColor: Print stxt(j).Caption;
                    ForeColor = Me.BackColor: Print String(T - Int(T / 2), "-");
                End If
                
                'Print stxt(j).Caption
            End With
        End If
    Next j
    
    '显示信号灯状态
    For j = 1 To Xhd_num
        If Val_Comp(Dev_XHD(j).Left + Dev_XHD(j).Width / 2, Dev_PIC(index).Left, Dev_PIC(index).Left + Dev_PIC(index).Width) And Val_Comp(Dev_XHD(j).Top + Dev_XHD(j).Height / 2, Dev_PIC(index).Top, Dev_PIC(index).Top + Dev_PIC(index).Height) Then
            F_ShowLc.DrawWidth = 1: F_ShowLc.ForeColor = QBColor(0)
            If Dev_XHD(j).Value = 0 Then        'xhd close
                F_ShowLc.FillColor = F_ShowLc.BackColor 'QBColor(12)
                Ellipse F_ShowLc.hdc, Dev_XHD(j).Left, Dev_XHD(j).Top, Dev_XHD(j).Left + Dev_XHD(j).Width, Dev_XHD(j).Top + Dev_XHD(j).Height
            Else                                        'xhd open
                F_ShowLc.FillColor = QBColor(10)
                Ellipse F_ShowLc.hdc, Dev_XHD(j).Left, Dev_XHD(j).Top, Dev_XHD(j).Left + Dev_XHD(j).Width, Dev_XHD(j).Top + Dev_XHD(j).Height
            End If
        End If
    Next j
    Dev_PIC(index).Last_Status = Dev_PIC(index).Value
    Dev_PIC(index).Last_Beizhu = Dev_PIC(index).Beizhu
End If
End Function

'******************************************************************************
'*  函数功能：运行指定胶料秤（CHENG_NAME：秤名；MACHINE：秤所属机号）（定时器用）*
'******************************************************************************
Public Function Run_JLC(Cheng_Name As String, Machine As Integer, Force_Show As Boolean)
Dim i, j As Integer
Dim Val_Chg As Boolean
For i = 1 To Pic_num
    If (Dev_PIC(i).Machine = Machine) And (Trim(Dev_PIC(i).Name) = Trim(Cheng_Name)) Then
        Val_Chg = IIf(Dev_PIC(i).Last_Status <> Dev_PIC(i).Value, True, False)
        If Val_Chg Or Force_Show Then
            '显示状态
            For j = 1 To Status_txt_num
                If Trim(Status_Txt(j).Relation_code) = Trim(Dev_PIC(i).Code) Then
                    stxt(j).Caption = str(Dev_PIC(i).Value) + "  " + (Dev_PIC(i).Beizhu)
                    With F_ShowLc
                        FontName = "Fixedsys" 'stxt(j).FontName
                        FontSize = stxt(j).FontSize
                        ForeColor = stxt(j).ForeColor
                        CurrentX = stxt(j).Left: CurrentY = stxt(j).Top
                        If Trim(stxt(j).Caption) <> "" Then
                            Print stxt(j).Caption
                        Else
                            Print "    "
                        End If
                    End With
                End If
            Next j
            
'            For T = 1 To Kj_num
'                If (Val(KJ(T).Tag) = i) Then
'                    KJ(T).Caption = CStr(Dev_PIC(i).Value)
'                    KJ(T).MaxWeight = Dev_PIC(i).Max
'                    KJ(T).weight = Dev_PIC(i).Value
'                    Exit For
'                End If
'            Next T
            Dev_PIC(i).Last_Status = Dev_PIC(i).Value
        End If
    End If
Next i
End Function

'******************************************************************************
'*  函数功能：运行指定（INDEX）的二次储斗（定时器用）                           *
'******************************************************************************
Public Function Run_ECCD(index As Integer, Force_Show As Boolean)
Dim i As Integer
Dim Bot_X, Bot_Y, High, K1, K2 As Double
Dim Wx, Wy, Ww, Wh, Fx, Fy, X1, Y1, X2, Y2, X3, Y3 As Long
Dim Fc, Bc As Long
Dim Val_Chg As Boolean
Dim Scale_X, Scale_Y As Double
Dim Zxy(0 To 3) As POINTAPI
Dim Tmp_X, Tmp_Y As Long
Dim Tmp_Val As Single
Fc = QBColor(5): Bc = QBColor(0)
Bot_X = 55# / 110#:     Bot_Y = 81# / 83#:     High = (81# - 27#) / 83#
'k1 = (16# - 44#) / (13# - 30#): k2 = (16# - 44#) / (49# - 30#)

Wx = Dev_PIC(index).Left '/ Window_Scale_X
Wy = Dev_PIC(index).Top '/ Window_Scale_Y
Ww = Dev_PIC(index).Width '/ Window_Scale_X
Wh = Dev_PIC(index).Height '/ Window_Scale_Y

Scale_X = Dev_PIC(index).Width / 110#
Scale_Y = Dev_PIC(index).Height / 83#
K1 = (1# - 55#) * Scale_Y / (27# - 81#) / Scale_X
K2 = (109# - 55#) * Scale_Y / (27# - 81#) / Scale_X

Val_Chg = IIf(Dev_PIC(index).Last_Status <> Dev_PIC(index).Value, True, False)
If Dev_PIC(index).Value > Dev_PIC(index).Max Then
    Tmp_Val = Dev_PIC(index).Max
Else
    If Dev_PIC(index).Value < 0 Then
        Tmp_Val = 0.01
    Else
        Tmp_Val = Dev_PIC(index).Value
    End If
End If
If Val_Chg Or Force_Show Then
    DrawWidth = 1
    'draw corner
    X1 = Wx + Ww * Bot_X:       Y1 = Wy + Wh * Bot_Y
    Y2 = Y1 - Wh * High:            X2 = (Y2 - Y1) / K1 + X1
    X3 = (Y2 - Y1) / K2 + X1:       Y3 = Y2
    
    Zxy(0).X = CLng(X2 * Window_Scale_X):    Zxy(0).Y = CLng(Y2 * Window_Scale_Y)
    Zxy(3).X = CLng(X3 * Window_Scale_X):    Zxy(3).Y = CLng(Y2 * Window_Scale_Y)
    
    Tmp_Y = CDbl(Y1) - Wh * High * CDbl(Tmp_Val) / CDbl(Dev_PIC(index).Max)
    'Tmp_Y = CDbl(Y1) - Wh * High * CDbl(Dev_pic(Index).Value) / CDbl(Dev_pic(Index).Max)
    Tmp_X = CDbl(Tmp_Y - Y1) / K1 + X1
    Zxy(1).X = CLng(Tmp_X * Window_Scale_X):    Zxy(1).Y = CLng(Tmp_Y * Window_Scale_Y)
    Tmp_X = CDbl(Tmp_Y - Y1) / K2 + X1
    Zxy(2).X = CLng(Tmp_X * Window_Scale_X):    Zxy(2).Y = CLng(Tmp_Y * Window_Scale_Y)

    '填充液面以上空间
    F_ShowLc.ForeColor = QBColor(0)
    F_ShowLc.FillStyle = 0
    F_ShowLc.FillColor = F_ShowLc.BackColor
    
    SetPolyFillMode F_ShowLc.hdc, 2
    Polygon F_ShowLc.hdc, Zxy(0), 4
    
    '填充液体
    Zxy(0).X = Zxy(1).X:     Zxy(0).Y = Zxy(1).Y
    Zxy(1).X = Zxy(2).X:     Zxy(1).Y = Zxy(2).Y
    Zxy(2).X = CLng(X1 * Window_Scale_X):    Zxy(2).Y = CLng(Y1 * Window_Scale_Y)
    
    F_ShowLc.ForeColor = QBColor(0)
    F_ShowLc.FillStyle = 0
    F_ShowLc.FillColor = Fc
    
    SetPolyFillMode F_ShowLc.hdc, 2
    Polygon F_ShowLc.hdc, Zxy(0), 3
    '--------------------------------------------------------------------
    '显示状态
    For j = 1 To Status_txt_num
        If Trim(Status_Txt(j).Relation_code) = Trim(Dev_PIC(index).Code) Then
            stxt(j).Caption = Format$(Dev_PIC(index).Value, "000.0")
            
            With F_ShowLc
                FontName = stxt(j).FontName:    FontSize = stxt(j).FontSize
                ForeColor = stxt(j).ForeColor
                If Dev_PIC(index).Value <= Dev_PIC(index).Max Then
                    If Dev_PIC(index).Value >= 0 Then
                        stxt(j).Caption = Trim(stxt(j).Caption) + "               "
                    Else
                        stxt(j).Caption = Trim(stxt(j).Caption) + "  异常"
                    End If
                Else
                    stxt(j).Caption = Trim(stxt(j).Caption) '+ " 溢出"
                End If
                CurrentX = stxt(j).Left: CurrentY = stxt(j).Top
                Print stxt(j).Caption
            End With
        End If
    Next j
    '显示信号灯状态
    For j = 1 To Xhd_num
        If Val_Comp(Dev_XHD(j).Left + Dev_XHD(j).Width / 2, Dev_PIC(index).Left, Dev_PIC(index).Left + Dev_PIC(index).Width) And Val_Comp(Dev_XHD(j).Top + Dev_XHD(j).Height / 2, Dev_PIC(index).Top, Dev_PIC(index).Top + Dev_PIC(index).Height) Then
            F_ShowLc.DrawWidth = 1: F_ShowLc.ForeColor = QBColor(0)
            If Dev_XHD(j).Value = 0 Then        'xhd close
                F_ShowLc.FillColor = F_ShowLc.BackColor 'QBColor(12)
                Ellipse F_ShowLc.hdc, Dev_XHD(j).Left, Dev_XHD(j).Top, Dev_XHD(j).Left + Dev_XHD(j).Width, Dev_XHD(j).Top + Dev_XHD(j).Height
            Else                                        'xhd open
                F_ShowLc.FillColor = QBColor(10)
                Ellipse F_ShowLc.hdc, Dev_XHD(j).Left, Dev_XHD(j).Top, Dev_XHD(j).Left + Dev_XHD(j).Width, Dev_XHD(j).Top + Dev_XHD(j).Height
            End If
        End If
    Next j
    Dev_PIC(index).Last_Status = Dev_PIC(index).Value
    Dev_PIC(index).Status = Tmp_Y
End If
End Function

'******************************************************************************
'*  函数功能：运行所有的阀门（定时器用）                                        *
'*  force_show = true    强制显示阀门状态
'******************************************************************************
Public Function Run_FaMen(Force_Show As Boolean)
Dim i As Integer
Dim Val_Chg As Boolean
Dim X1, Y1, W1, H1, X2, Y2, W2, H2 As Long
    For i = 1 To Pic_num
        If (Dev_PIC(i).Style = 1) Then    '阀门
            Val_Chg = IIf(Dev_PIC(i).Last_Status <> Dev_PIC(i).Value, True, False)
            Select Case Dev_PIC(i).Action
                Case 0  'Run famen
                    If ((Dev_PIC(i).Enabled = True) And Val_Chg) Or Force_Show Then
        '                X1 = Window_Scale_X * (Dev_pic(i).Left + Dev_pic(i).Width * Fm_hole(Dev_pic(i).Position).scale_x1)
        '                Y1 = Window_Scale_Y * (Dev_pic(i).Top + Dev_pic(i).Height * Fm_hole(Dev_pic(i).Position).scale_y1)
        '                W1 = Window_Scale_X * (Dev_pic(i).Width * Fm_hole(Dev_pic(i).Position).scale_w1)
        '                H1 = Window_Scale_Y * (Dev_pic(i).Height * Fm_hole(Dev_pic(i).Position).scale_h1)
                        
        '                X2 = Window_Scale_X * (Dev_pic(i).Left + Dev_pic(i).Width * Fm_hole(Dev_pic(i).Position).scale_x2)
        '                Y2 = Window_Scale_Y * (Dev_pic(i).Top + Dev_pic(i).Height * Fm_hole(Dev_pic(i).Position).scale_y2)
        '                W2 = Window_Scale_X * (Dev_pic(i).Width * Fm_hole(Dev_pic(i).Position).scale_w2)
        '                H2 = Window_Scale_Y * (Dev_pic(i).Height * Fm_hole(Dev_pic(i).Position).scale_h2)
                        
                        X1 = (Dev_PIC(i).Left + Dev_PIC(i).Width * Fm_hole(Dev_PIC(i).Position).scale_x1)
                        Y1 = (Dev_PIC(i).Top + Dev_PIC(i).Height * Fm_hole(Dev_PIC(i).Position).scale_y1)
                        W1 = (Dev_PIC(i).Width * Fm_hole(Dev_PIC(i).Position).scale_w1)
                        H1 = (Dev_PIC(i).Height * Fm_hole(Dev_PIC(i).Position).scale_h1)
                        
                        X2 = (Dev_PIC(i).Left + Dev_PIC(i).Width * Fm_hole(Dev_PIC(i).Position).scale_x2)
                        Y2 = (Dev_PIC(i).Top + Dev_PIC(i).Height * Fm_hole(Dev_PIC(i).Position).scale_y2)
                        W2 = (Dev_PIC(i).Width * Fm_hole(Dev_PIC(i).Position).scale_w2)
                        H2 = (Dev_PIC(i).Height * Fm_hole(Dev_PIC(i).Position).scale_h2)
                        
                        If Dev_PIC(i).Value = 0 Then
                            'xhd(i).FillColor = QBColor(0)          'light red
                            F_ShowLc.FillColor = F_ShowLc.BackColor 'QBColor(2)
                            F_ShowLc.DrawWidth = 1: F_ShowLc.ForeColor = QBColor(0)
                            F_ShowLc.Circle (X1 + W1 / 2, Y1 + H1 / 2), IIf(H1 > W1, H1 / 2, W1 / 2), , , , H1 / W1
                            F_ShowLc.Circle (X2 + W2 / 2, Y2 + H2 / 2), IIf(H2 > W2, H2 / 2, W2 / 2), , , , H2 / W2
                            'Ellipse F_ShowLc.hdc, X1, Y1, X1 + W1, Y1 + H1
                            'Ellipse F_ShowLc.hdc, X2, Y2, X2 + W2, Y2 + H2
                            
                            For j = 1 To Status_txt_num
                                If Trim(Status_Txt(j).Relation_code) = Trim(Dev_PIC(i).Code) Then
                                    stxt(j).Caption = "关"
                                    With F_ShowLc
                                        CurrentX = stxt(j).Left: CurrentY = stxt(j).Top
                                        FontName = stxt(j).FontName:    FontSize = stxt(j).FontSize
                                        ForeColor = stxt(j).ForeColor
                                        Print stxt(j).Caption
                                    End With
                                End If
                            Next j
                        Else
                            'xhd(i).FillColor = QBColor(10)          'light green
                            F_ShowLc.FillColor = QBColor(13)    'light magenta
                            F_ShowLc.DrawWidth = 1: F_ShowLc.ForeColor = QBColor(0)
                            F_ShowLc.Circle (X1 + W1 / 2, Y1 + H1 / 2), IIf(H1 > W1, H1 / 2, W1 / 2), , , , H1 / W1
                            F_ShowLc.Circle (X2 + W2 / 2, Y2 + H2 / 2), IIf(H2 > W2, H2 / 2, W2 / 2), , , , H2 / W2
        '                    Ellipse F_ShowLc.hdc, X1, Y1, X1 + W1, Y1 + H1
        '                    Ellipse F_ShowLc.hdc, X2, Y2, X2 + W2, Y2 + H2
                            
                        End If
                        Dev_PIC(i).Last_Status = Dev_PIC(i).Value
                    End If
                Case 1  'Hide famen
                    If Val_Chg Then
                        Dev_PIC(i).Last_Status = Dev_PIC(i).Value
                    End If
            End Select
        End If  'Is famen
    Next i
End Function

'****************************************************************************
'*  函数功能：运行所有的碳黑秤、油料秤（定时器用）                            *
'****************************************************************************
Public Function Run_THC_YLCS(Force_Show As Boolean)
Dim i As Integer
For i = 1 To Pic_num
    If (Dev_PIC(i).Style = 7) And (Dev_PIC(i).Max > 0) Then          '碳黑秤,油料秤
            Run_THC_YLC i, Force_Show
    End If
Next i
End Function

'****************************************************************************
'*  函数功能：运行所有的胶料秤（定时器用）                            *
'****************************************************************************
Public Function Run_JLCS(Force_Show As Boolean)
Dim i As Integer
For i = 1 To Pic_num
    If (Dev_PIC(i).Style = 6) And (Dev_PIC(i).Max > 0) Then     '胶料秤
        Run_JLC Trim(Dev_PIC(i).Name), Dev_PIC(i).Machine, Force_Show
    End If
Next i
End Function

'****************************************************************************
'*  函数功能：运行所有的二次储斗（定时器用）                                  *
'****************************************************************************
Public Function Run_ECCDS(Force_Show As Boolean)
Dim i As Integer
For i = 1 To Pic_num
    If (Dev_PIC(i).Style = 18) And (Dev_PIC(i).Max > 0) Then          '二次储斗
        Run_ECCD i, Force_Show
    End If
Next i
End Function

'****************************************************************************
'*  函数功能：运行所有的信号灯（定时器用）                                    *
'*  froce_show = true  强制显示信号灯状态                                    *
'****************************************************************************
Public Function Run_XHD(Force_Show As Boolean)
Dim i, j As Integer
Dim Val_Chg As Boolean
    For i = 1 To Xhd_num
        Val_Chg = IIf(Dev_XHD(i).Last_Status <> Dev_XHD(i).Value, True, False)
        Select Case Dev_XHD(i).Action
            Case 0  'Run xhd
                If ((Dev_XHD(i).Enabled = True) And Val_Chg) Or Force_Show Then
                    If Dev_XHD(i).Value = 0 Then
                        'xhd(i).FillColor = QBColor(12)          'light red
                        F_ShowLc.FillColor = QBColor(2)         'F_ShowLc.BackColor
                        F_ShowLc.DrawWidth = 1: F_ShowLc.ForeColor = QBColor(0)
                        Ellipse F_ShowLc.hdc, Dev_XHD(i).Left, Dev_XHD(i).Top, Dev_XHD(i).Left + Dev_XHD(i).Width, Dev_XHD(i).Top + Dev_XHD(i).Height
                        For j = 1 To Status_txt_num
                            If Trim(Status_Txt(j).Relation_code) = Trim(Dev_XHD(i).Code) Then
                                stxt(j).Caption = "关"
                                With F_ShowLc
                                    CurrentX = stxt(j).Left: CurrentY = stxt(j).Top
                                    FontName = stxt(j).FontName:    FontSize = stxt(j).FontSize
                                    ForeColor = stxt(j).ForeColor
                                    Print stxt(j).Caption
                                End With
                            End If
                        Next j
                    Else
                        'xhd(i).FillColor = QBColor(10)          'light green
                        F_ShowLc.FillColor = QBColor(10)
                        F_ShowLc.DrawWidth = 1: F_ShowLc.ForeColor = QBColor(0)
                        Ellipse F_ShowLc.hdc, Dev_XHD(i).Left, Dev_XHD(i).Top, Dev_XHD(i).Left + Dev_XHD(i).Width, Dev_XHD(i).Top + Dev_XHD(i).Height
                    End If
                    Dev_XHD(i).Last_Status = Dev_XHD(i).Value
                End If
            Case 1 'Hide xhd
                If Val_Chg Then
                    Dev_XHD(i).Last_Status = Dev_XHD(i).Value
                End If
        End Select
    Next i
End Function

'****************************************************************************
'*  函数功能：获取管道数据                                                   *
'* STYLE：      管道类型
'* COLOUR：     管道颜色
'* WIDE：       管道宽度
'* OBJ_COLOR：  流体宽度
'****************************************************************************
Public Sub Get_GD_Property(Style As Integer, Colour As Long, Wide As Integer, Nj_Width As Integer, Obj_color As Long)
Select Case Style
    Case 11
            Colour = Gqg.Color: Wide = Gqg.Width: Nj_Width = Gqg.Neijing_width: Obj_color = Gqg.Obj_color
    Case 12
            Colour = Gyg.Color: Wide = Gyg.Width: Nj_Width = Gyg.Neijing_width: Obj_color = Gyg.Obj_color
    Case 13
            Colour = Glg.Color: Wide = Glg.Width: Nj_Width = Glg.Neijing_width: Obj_color = Glg.Obj_color
End Select
End Sub

'****************************************************************************
'*  函数功能：运行指定（INDEX）的管道                                        *
'****************************************************************************
Public Function Run_Zline(index As Integer)
Dim i, W As Integer
Dim Nj_W As Integer
Dim C As Long
Dim Nj_C As Long
Call Get_GD_Property(Dev_ZLine(index).Style, C, W, Nj_W, Nj_C)
If DrawWidth <> Nj_W Then
    DrawWidth = Nj_W
End If
For i = 0 To Dev_ZLine(index).Zdot_num - 2
    Draw_Line F_ShowLc, Dev_ZLine(index).Zxy(i, 0), Dev_ZLine(index).Zxy(i, 1), Dev_ZLine(index).Zxy(i + 1, 0), Dev_ZLine(index).Zxy(i + 1, 1), Dev_ZLine(index).Cur_Offset, C, Nj_C
Next i
End Function

'****************************************************************************
'*  函数功能：运行所有的管道                                                 *
'*  由管道的Obj_passing属性决定管道是否运动                                  *
'****************************************************************************
Function Run_Zlines()
Dim i As Integer
'Dash_line_offset = Dash_line_offset + 2
'If Dash_line_offset > Dls.Xd_len + Dls.Space_len - 2 Then
'    Dash_line_offset = 0
'End If
    For i = 0 To Zline_num - 1
        Dev_ZLine(i).Cur_Offset = Dev_ZLine(i).Cur_Offset + Dev_ZLine(i).Speed
        If Dev_ZLine(i).Cur_Offset > (Dls.Xd_len + Dls.Space_len - Dev_ZLine(i).Speed) Then
            Dev_ZLine(i).Cur_Offset = 0
        End If
        
        If Dev_ZLine(i).Obj_Passing Then
            If Dev_ZLine(i).Visible = False Then
                Hide_Zline i, False
            End If
            Run_Zline (i)
        Else
            If Dev_ZLine(i).Visible = True Then
                Hide_Zline i, True
            End If
        End If
    Next i
End Function

'****************************************************************************
'*  函数功能：隐藏指定（INDEX）的管道                                         *
'*  折线状态设置 T=true 隐藏  T=false 显示                                   *
'****************************************************************************
Function Hide_Zline(index As Integer, T As Boolean)
Dim i, Wide As Integer
Dim Colour As Long
Dim Nj_C As Long
Dim Nj_Wide As Integer
    Call Get_GD_Property(Dev_ZLine(index).Style, Colour, Wide, Nj_Wide, Nj_C)
    F_ShowLc.CurrentX = Dev_ZLine(index).Zxy(0, 0)
    F_ShowLc.CurrentY = Dev_ZLine(index).Zxy(0, 1)
    For i = 1 To Dev_ZLine(index).Zdot_num - 1
        F_ShowLc.DrawWidth = Wide: F_ShowLc.DrawStyle = 0
        Line -(Dev_ZLine(index).Zxy(i, 0), Dev_ZLine(index).Zxy(i, 1)), Colour
    Next i
    Dev_ZLine(index).Last_Status = 0
    Dev_ZLine(index).Visible = IIf(T = True, False, True)
End Function

'****************************************************************************
'*  函数功能：由elect_receive_input.data设置阀门、信号灯的value,can_comein    *
'*            ,can_goout,obj_passing属性                                     *
'****************************************************************************
Public Function Set_FaMens_XHDS_Status()
Dim i, j, k As Integer
Dim X As elect_send_type
Dim Y As elect_send_type
        Tmp = read_elec_input(Y)                        '读信号灯数据
'        If Elec_Run!Change_Command.Tag = "N" Then
                Tmp = read_elec_output(X)               '读阀门数据
'        End If
'设置阀门
For i = 1 To Pic_num
    If Dev_PIC(i).Enabled And ((Dev_PIC(i).Style = 1) Or (Dev_PIC(i).Style = 5)) Then
        For j = 0 To X.Length
            If Trim(Dev_PIC(i).Name) = Trim(X.Note_Name(j)) Then
                Dev_PIC(i).Value = X.data(j)
                If (Dev_PIC(i).Value <> 0) Then    '启动
                    Dev_PIC(i).Can_Comein = True
                    Dev_PIC(i).Can_Goout = True
                Else                                                                                '不启动
                    Dev_PIC(i).Can_Comein = False
                    Dev_PIC(i).Can_Goout = False
                End If
                GoTo Jump_Out_1
            End If
        Next j
        For j = 0 To Y.Length
            If Trim(Dev_PIC(i).Name) = Trim(Y.Note_Name(j)) Then
                Dev_PIC(i).Value = Y.data(j)
                If (Dev_PIC(i).Value <> 0) Then    '启动
                    Dev_PIC(i).Can_Comein = True
                    Dev_PIC(i).Can_Goout = True
                Else                                                                                '不启动
                    Dev_PIC(i).Can_Comein = False
                    Dev_PIC(i).Can_Goout = False
                End If
                GoTo Jump_Out_1
            End If
        Next j
    End If
Jump_Out_1:
Next i
'设置信号灯
For i = 1 To Xhd_num
    If Dev_XHD(i).Enabled Then
        For j = 0 To Y.Length
            If Trim(Dev_XHD(i).Name) = Trim(Y.Note_Name(j)) Then
                Dev_XHD(i).Value = Y.data(j)
                GoTo Jump_Out_2
            End If
        Next j
        For j = 0 To X.Length
            If Trim(Dev_XHD(i).Name) = Trim(X.Note_Name(j)) Then
                Dev_XHD(i).Value = X.data(j)
                GoTo Jump_Out_2
            End If
        Next j
    End If
Jump_Out_2:
Next i
End Function

'****************************************************************************
'*  函数功能：设置当前使用流体网络算法得到的路径中的设备运行参数               *
'****************************************************************************
Public Function Luo_Set_Road()
Dim i, j, sp, ep, sr, er, Tmp_Pic As Integer
Dim Can_Passing As Boolean

sr = 0: er = 0
sp = Dev_ZLine(Road.ZIndex(sr)).Start_Pic_Index
ep = Dev_ZLine(Road.ZIndex(sr)).End_Pic_Index
If Road.Length <= 0 Then
    Exit Function
End If
While er <= (Road.Length - 1)
    '查找不是阀门或隐含阀门的pic
    While ((Dev_PIC(ep).Style = 1) Or (Dev_PIC(ep).Style = 5)) And (er < (Road.Length - 1))
        er = er + 1
        ep = Dev_ZLine(Road.ZIndex(er)).End_Pic_Index
    Wend
    '判断两PIC之间是否为通路
    If (Dev_PIC(sp).Can_Goout And Dev_PIC(sp).Obj_Passing) And Dev_PIC(ep).Can_Comein Then
        Can_Passing = True
        For i = sr To er
            Tmp_Pic = Dev_ZLine(Road.ZIndex(i)).End_Pic_Index
'            If Dev_pic(Tmp_Pic).Can_comein And Dev_pic(Tmp_Pic).Can_goout Then '开
            If Dev_PIC(Tmp_Pic).Value <> 0 Then
            Else
                Can_Passing = False
                Exit For
            End If
        Next i
        
'        For i = sr To er
'            Tmp_pic = Dev_zline(road.zindex(i)).End_pic_index
'            If Dev_pic(Tmp_pic).Can_comein And Dev_pic(Tmp_pic).Can_goout Then '开
'            If dev_pic(dev_zline(road.zindex(i)).end_pic_index).Value <> 0 Then '开
'                Dev_pic(Dev_zline(road.zindex(i)).End_pic_index).Can_comein = True
'                Dev_pic(Dev_zline(road.zindex(i)).End_pic_index).Can_goout = True
'                Dev_pic(Dev_zline(road.zindex(i)).End_pic_index).Obj_passing = True
'            Else
'                Can_passing = False
'                Exit For
'            End If
'        Next i
    Else
        Can_Passing = False
    End If
    
    '通路
    If Can_Passing Then
        For i = sr To er
            Dev_PIC(Dev_ZLine(Road.ZIndex(i)).End_Pic_Index).Obj_Passing = True
            Zline_obj_passing(Road.ZIndex(i)) = True
        Next i
        Dev_PIC(ep).Can_Comein = True
        Dev_PIC(ep).Can_Goout = True
        Dev_PIC(ep).Obj_Passing = True
        
        'Change Zline's Speed based on dev_pic's beizhu
        If Dev_PIC(ep).Style = 7 Then       '碳黑秤，油料秤
            Select Case Trim(Dev_PIC(ep).Beizhu)
                Case "快加"
                    For j = sr To er
                        Dev_ZLine(Road.ZIndex(j)).Speed = LS_FAST2
                    Next j
                Case "慢加"
                    For j = sr To er
                        Dev_ZLine(Road.ZIndex(j)).Speed = LS_SLOW
                    Next j
                Case "点动"
                    For j = sr To er
                        Dev_ZLine(Road.ZIndex(j)).Speed = 0.5
                    Next j
                Case "等待"
                    For j = sr To er
                        Dev_ZLine(Road.ZIndex(j)).Speed = 0
                    Next j
                Case Else
                    For j = sr To er
                        Dev_ZLine(Road.ZIndex(j)).Speed = LS_NORMAL
                    Next j
            End Select
        End If
    End If
    
    sr = er + 1
    'If sr <= (road.length - 1) Then
        er = sr
        sp = Dev_ZLine(Road.ZIndex(sr)).Start_Pic_Index
        ep = Dev_ZLine(Road.ZIndex(sr)).End_Pic_Index
    'End If
'    tmp = DoEvents
'Loop 'Until sr > (road.length - 1)
Wend
End Function

'****************************************************************************
'  函数功能：使能指定（S）的设备                                              *
'****************************************************************************
Public Function Set_Pic_Enabled(s As Integer, E As Boolean)
Dim i As Integer
For i = 1 To Pic_num
    If Dev_PIC(i).Style = s Then
        If E = True Then
            Dev_PIC(i).Can_Comein = True
            Dev_PIC(i).Can_Goout = True
            Dev_PIC(i).Obj_Passing = True
            If (Dev_PIC(i).Style = 6) Or (Dev_PIC(i).Style = 7) Then
                Dev_PIC(i).Value = 0
            Else
                Dev_PIC(i).Value = 1
            End If
        Else
            Dev_PIC(i).Can_Comein = False
            Dev_PIC(i).Can_Goout = False
            Dev_PIC(i).Obj_Passing = False
            Dev_PIC(i).Value = 0
        End If
    End If
Next i
End Function

'****************************************************************************
'* 函数功能：设置所有碳黑秤、油料秤的运行参数                                 *
'****************************************************************************
Public Function Set_THC_YLCS_Status()
Dim i, j, Tmp, Tmp_Flag As Integer
Dim Tmp_Max, Tmp_Wei As Single
For i = 1 To Pic_num
    If Dev_PIC(i).Style = 7 Then    '碳黑秤,油料秤
        '获取指定碳黑秤或油料秤的运行参数
        Tmp = read_ban(Dev_PIC(i).Machine, CInt(Left(Trim(Dev_PIC(i).Name), 1)), Dev_PIC(i).Max, Tmp_Wei, Tmp_Flag)
        If Dev_PIC(i).Enabled = True Then
            Tmp_Wei = Int(Tmp_Wei * 100 + 0.5) / 100
            Dev_PIC(i).Value = Tmp_Wei
            Select Case Tmp_Flag
                Case 1
                    Dev_PIC(i).Beizhu = ""
                Case 2
                    Dev_PIC(i).Beizhu = "快加"
                Case 3
                    Dev_PIC(i).Beizhu = "慢加"
                Case 4
                    Dev_PIC(i).Beizhu = "点动"
                Case 5
                    Dev_PIC(i).Beizhu = "等待"
                Case Else
            End Select
            If Tmp_Wei > CSng(Dev_PIC(i).Max_Size) Then
                Dev_PIC(i).Beizhu = "溢出"
            Else
                If Tmp_Wei < 0 Then
                    Dev_PIC(i).Beizhu = "异常"
                End If
            End If
           
            If (Dev_PIC(i).Value <> 0) And Dev_PIC(i).Enabled Then
                Dev_PIC(i).Can_Comein = True
                Dev_PIC(i).Can_Goout = True
                Dev_PIC(i).Obj_Passing = True
            Else
                Dev_PIC(i).Can_Comein = False
                Dev_PIC(i).Can_Goout = False
                Dev_PIC(i).Obj_Passing = False
            End If
        End If  'enabled=true
    End If      'style = 7
Next i
End Function

'****************************************************************************
'* 函数功能：设置所有胶料秤的运行参数                                         *
'****************************************************************************
Public Function Set_JLCS_Status()
Dim Tmp, Tmp_Flag As Integer
Dim Tmp_Max, Tmp_Wei As Single
For i = 1 To Pic_num
    If Dev_PIC(i).Style = 6 Then    '胶料秤
        '获取指定胶料秤的运行参数
        Tmp = read_ban(Dev_PIC(i).Machine, CInt(Left(Trim(Dev_PIC(i).Name), 1)), Tmp_Max, Tmp_Wei, Tmp_Flag)
        If Dev_PIC(i).Enabled = True Then
            Dev_PIC(i).Max = IIf(Tmp_Max > 0, Tmp_Max, 0)
            Dev_PIC(i).Value = Int(Tmp_Wei * 10 + 0.5) / 10
            Select Case Tmp_Flag
                Case 1
                    Dev_PIC(i).Beizhu = "块胶"
                Case 2
                    Dev_PIC(i).Beizhu = "片胶"
                Case Else
                    Dev_PIC(i).Beizhu = ""
            End Select
            
            If Dev_PIC(i).Value > 0 Then
                Dev_PIC(i).Can_Comein = True
                Dev_PIC(i).Can_Goout = True
                Dev_PIC(i).Obj_Passing = True
            Else
                Dev_PIC(i).Can_Comein = False
                Dev_PIC(i).Can_Goout = False
                Dev_PIC(i).Obj_Passing = False
            End If
        End If
    End If
Next i
End Function

'****************************************************************************
'* 函数功能：设置所有二次储斗的运行参数                                       *
'* 参数要求：二次储斗尾部数字与秤的尾部数字相对应                             *
'****************************************************************************
Public Function Set_ECCDS_Status()
Dim index, j, k, L As Integer
Dim Eccd_can_comein, Eccd_can_goout As Boolean
Dim Tmp_Val As Single
For index = 1 To Pic_num
    If (Dev_PIC(index).Style = 18) And Dev_PIC(index).Enabled Then     '二次储斗
        '查找与本二次储斗相对应的秤并设置其min,max
        For k = 1 To Pic_num
            If (Dev_PIC(k).Style = 7) And (Dev_PIC(k).Machine = Dev_PIC(index).Machine) And (Left(Trim(Dev_PIC(k).Name), 1) = Left(Trim(Dev_PIC(index).Name), 1)) Then
                Dev_PIC(index).Min = Dev_PIC(k).Min
                Dev_PIC(index).Max = Dev_PIC(k).Max
                Exit For
            End If
        Next k
        Eccd_can_comein = False
        For j = 1 To Pic_num
            If (Arcs(j, index) <> -1) Then
                 If Dev_PIC(Dev_ZLine(Arcs(j, index)).Start_Pic_Index).Can_Goout Then
                    Eccd_can_comein = True
                    Exit For
                End If
            End If
        Next j
        Dev_PIC(index).Can_Comein = Eccd_can_comein
        Eccd_can_goout = False
        For j = 1 To Pic_num
            If Arcs(index, j) <> -1 Then
                 If Dev_PIC(Dev_ZLine(Arcs(index, j)).End_Pic_Index).Can_Comein Then
                    Eccd_can_goout = True
                    Exit For
                End If
            End If
        Next j
        Dev_PIC(index).Can_Goout = Eccd_can_goout
        
        '若eccd的前驱图象can_goout=true obj_passing=true,后继图象can_comein=false则表示正在装入碳黑
        If Dev_PIC(index).Can_Comein Then
            If (Dev_PIC(k).Last_Status - Dev_PIC(k).Value) > 0 Then   '油料秤(碳黑秤)正在流出液体
                Tmp_Val = Dev_PIC(index).Value + (Dev_PIC(k).Last_Status - Dev_PIC(k).Value)
            Else
                Tmp_Val = 0.01 'Dev_pic(index).Value
            End If
            Dev_PIC(index).Value = Tmp_Val
        End If
        
        '若前驱图象can_goout=false,后继图象can_comein=true则表示碳黑正在流出
        'If (Not Dev_pic(Index).Can_comein) And Dev_pic(Index).Can_goout Then
        If Dev_PIC(index).Can_Goout Then
            Tmp_Val = Dev_PIC(index).Value - Dev_PIC(index).Max / 30
            If Tmp_Val > 0 Then
                Dev_PIC(index).Value = Tmp_Val
            Else
                Dev_PIC(index).Value = 0.01
            End If
        End If
        
        Dev_PIC(index).Obj_Passing = IIf(Dev_PIC(index).Value > 0, True, False)
    End If          'style=18
Next index
End Function

'****************************************************************************
'* 函数功能：设置所有日储斗的运行参数                                         *
'****************************************************************************
Public Function Set_RCDS_Status()
Dim i, Tmp As Integer
Dim RCD_In_Flag As Integer
Dim RCD_Out_Flag As Integer
Dim RCD_Name As Integer
For i = 1 To Pic_num
    If (Dev_PIC(i).Style = 4) Then    '日储斗
        RCD_Name = CInt(Left(Trim(Dev_PIC(i).Name), 1))
        '获取指定日储斗的进口运行参数
        Tmp = read_Tan_dou(Dev_PIC(i).Machine, RCD_Name, RCD_In_Flag)
        '获取指定日储斗的出口运行参数
'        Tmp = read_ri_chu_dou(Dev_PIC(i).Machine, RCD_Name, RCD_Out_Flag)
        If Dev_PIC(i).Enabled Then
            Dev_PIC(i).Obj_Passing = True
            If RCD_In_Flag = 1 Then
                Dev_PIC(i).Can_Comein = True
            Else
                Dev_PIC(i).Can_Comein = False
            End If
'            If RCD_Out_Flag = 1 Then
                Dev_PIC(i).Can_Goout = True
'            Else
'                Dev_PIC(i).Can_Goout = False
'            End If
        Else
            Dev_PIC(i).Can_Comein = False
            Dev_PIC(i).Can_Goout = False
            Dev_PIC(i).Obj_Passing = False
        End If
    End If
Next i
End Function

'****************************************************************************
'* 函数功能：设置所有油料斗的运行参数                                         *
'****************************************************************************
Public Function set_ylds_status()
Dim i, Tmp, Flag As Integer
For i = 1 To Pic_num
    If (Dev_PIC(i).Style = 19) And (Left(Trim(Dev_PIC(i).Name), 1) >= "0" And Left(Trim(Dev_PIC(i).Name), 1) <= "9") Then '油料斗
        '获取指定油料斗的运行参数
        Tmp = read_You_dou(Dev_PIC(i).Machine, CInt(Left(Trim(Dev_PIC(i).Name), 1)), Flag)
        If Dev_PIC(i).Enabled And (Flag = 1) Then     '开始选斗
            Dev_PIC(i).Can_Comein = True
            Dev_PIC(i).Can_Goout = True
            Dev_PIC(i).Obj_Passing = True
        End If
        If Dev_PIC(i).Enabled And (Flag = 0) Then       '停选
            Dev_PIC(i).Can_Comein = False
            Dev_PIC(i).Can_Goout = False
            Dev_PIC(i).Obj_Passing = False
        End If
    End If
Next i
End Function

'****************************************************************************
'*  函数功能：将折线数据转化为流体网络用的邻接矩阵数据                         *
'*  限制：作图时两个顶点之间不能有两条以上的相同方向的折线                     *
'****************************************************************************
Public Function DFS_Zline_To_Arcs()
Dim i, j As Integer
For i = 1 To Pic_num
    For j = 1 To Pic_num
        Arcs(i, j) = -1
    Next j
Next i
For i = 0 To Zline_num - 1
    If Dev_ZLine(i).Style <> 11 Then   '不是供汽管
        Arcs(Dev_ZLine(i).Start_Pic_Index, Dev_ZLine(i).End_Pic_Index) = i
    End If
Next i
End Function

Private Sub Adderr_Click()
Dim Y As String
Y = InputBox("error:", "", "100")
If Trim(Y) <> "" Then
    Call write_error(CLng(Y), "Temp test!")
End If
End Sub

Private Sub DelErr_Click()
Dim Y As String
Y = InputBox("error:", "", "100")
If Trim(Y) <> "" Then
    Call kill_error(CLng(Y))
End If

End Sub

Private Sub ErrorGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim Error_Code, Error_Prompt As String
If F_ShowLc.Visible Then
ErrorGrid.SetFocus
Select Case KeyCode
    Case vbKeyReturn
        Call m_o_errorhelp_Click
    Case vbKeyF3
        If Not F_Error_Help.Visible Then
            ErrorGrid.row = ErrorGrid.TopRow
            ErrorGrid.col = 1
            Error_Code = Trim(ErrorGrid.Text)
            ErrorGrid.col = 2
            Error_Prompt = Trim(ErrorGrid.Text)
            
            ErrorData.Recordset.index = "error_idx"
            ErrorData.Recordset.Seek "=", Error_Code
            If Not ErrorData.Recordset.NoMatch Then
                F_Error_Help.Label2.Caption = Error_Prompt
                F_Error_Help.Text1.Text = ErrorData.Recordset!China_Help
                F_Error_Help.Show vbModal
                Reading_ErrorHelp = True
            Else
                F_Error_Help.Label2.Caption = Error_Prompt
                F_Error_Help.Text1.Text = ""
            End If
        End If
    Case vbKeyF2
        For i = 0 To Cur_Errors.Length - 1
            Call kill_error(CLng(Cur_Errors.Code(i)))
        Next i
End Select
End If
End Sub

Private Sub Form_Activate()
Dim i As Integer
For i = 0 To Zline_num - 1
    If Left(Dev_ZLine(i).Code, 1) <> "Q" Then
        Dev_ZLine(i).Can_Comein = True
        Dev_ZLine(i).Can_Goout = True
    End If
Next i
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
If F_ShowLc.Visible And ErrorGrid.Visible Then
    ErrorGrid.SetFocus
    Select Case KeyCode
        Case vbKeyRight
            SendKeys "{RIGHT}"
        Case vbKeyLeft
            SendKeys "{LEFT}"
        Case vbKeyDown
            SendKeys "{DOWN}"
        Case vbKeyUp
            SendKeys "{UP}"
        Case vbKeyReturn
            'SendKeys "{ENTER}"
    End Select
    Call ErrorGrid_KeyDown(KeyCode, Shift)
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
Dim i As Integer
F_ShowLc.Left = 0
F_ShowLc.Top = 0
Timer1.Enabled = False
Timer2.Enabled = False
Call Add_Win(Me.Hwnd)
For i = 0 To Zline_num - 1
    Dev_ZLine(i).Can_Comein = True
    Dev_ZLine(i).Can_Goout = True
Next i
'  change by zgh
'  add  make_mathine
If Make_Mathine <> 0 Then
    ErrorGrid.Top = Label(7).Top
End If
Call Main(Make_Mathine)
Timer2.Enabled = True
Timer1.Enabled = True
Call SetThreadPriority(GetCurrentThread(), 8)
End Sub

'显示当前的系统图
Private Sub Form_Paint()
Dim i, j, k, Wide As Integer
Dim Nj_Wide As Integer
Dim Colour As Long
Dim Nj_C As Long
Dim count As Integer

F_ShowLc.MousePointer = 11

MAIN_MDI.Process_Msg.Visible = True
MAIN_MDI.ProcessBar.Visible = True

MAIN_MDI.Process_Msg.Caption = "放置设备"
MAIN_MDI.ProcessBar.Min = 0
MAIN_MDI.ProcessBar.Max = IIf(Pic_num <= 0, 1, Pic_num)
MAIN_MDI.ProcessBar.Value = 0

ComData.Refresh
'F_ShowLc.ComData.Recordset.index = "Style_Idx"
'F_ShowLc.ComData.Recordset.Seek "=", Dev_pic(index).Style
count = 0
On Error Resume Next
ComData.Recordset.MoveFirst
While Not ComData.Recordset.EOF
    For i = 1 To Pic_num
        If Dev_PIC(i).Style = ComData.Recordset!Style_id Then
            DoEvents
            ReadDb_Draw_Pic (i)
            count = count + 1
            MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
            If count > Pic_num Then
                GoTo WExit
            End If
        End If
    Next i
    ComData.Recordset.MoveNext
Wend
WExit:
'If Not ComData.Recordset.NoMatch Then
'End If

'For i = 1 To Pic_num
'    If Dev_pic(i).Action <> 1 Then '不是隐藏的图象
'        DoEvents
'        ReadDb_Draw_Pic (i)
'        MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
'    End If
'Next i

MAIN_MDI.Process_Msg.Caption = "放置管道"
MAIN_MDI.ProcessBar.Min = 0
MAIN_MDI.ProcessBar.Max = IIf(Zline_num <= 0, 1, Zline_num)
MAIN_MDI.ProcessBar.Value = 0
For i = 0 To Zline_num - 1
    Call Get_GD_Property(Dev_ZLine(i).Style, Colour, Wide, Nj_Wide, Nj_C)
    F_ShowLc.CurrentX = Dev_ZLine(i).Zxy(0, 0)
    F_ShowLc.CurrentY = Dev_ZLine(i).Zxy(0, 1)
    For j = 1 To Dev_ZLine(i).Zdot_num - 1
        F_ShowLc.DrawWidth = Wide
        F_ShowLc.DrawStyle = 0
        F_ShowLc.Line -(Dev_ZLine(i).Zxy(j, 0), Dev_ZLine(i).Zxy(j, 1)), Colour
    Next j
    DoEvents
    MAIN_MDI.ProcessBar.Value = MAIN_MDI.ProcessBar.Value + 1
Next i

Run_FaMen (True)
Run_XHD (True)
Run_THC_YLCS (True)
Run_LJLS (True)

F_ShowLc.FontTransparent = True
  For i = 1 To Txt_num
    F_ShowLc.ForeColor = txt(i).ForeColor
    F_ShowLc.Font = txt(i).Font
    F_ShowLc.FontName = txt(i).FontName
    F_ShowLc.FontSize = txt(i).FontSize
    Me.CurrentX = txt(i).Left
    Me.CurrentY = txt(i).Top
    
    F_ShowLc.Print txt(i).Caption
'    F_ShowLc!txt(i).Visible = True
'    F_ShowLc!txt(i).ZOrder 0
  Next i
F_ShowLc.FontTransparent = False

'----------------MENU-----------------------------
If China_English = True Then
    m_window.Caption = "窗口(&W)"
    m_w_liucheng.Caption = "关闭演示窗口"
    m_option.Caption = "设置(&O)"
    m_o_zoom.Caption = "缩放窗口"
    m_o_guandao.Caption = "管道属性"
    AddErr.Caption = "增加错误"
    DelErr.Caption = "删除错误"
    ' change by zgh 1998
    If Make_Mathine <> 0 Then
        Label(0).Caption = "车数"
        Label(1).Caption = "运行车号"
        Label(2).Caption = "段号"
        Label(3).Caption = "温度"
        Label(4).Caption = "段时间"
        Label(5).Caption = "总时间"
        Label(6).Caption = "能量"

    Else
        Label(0).Caption = "机组号"
        Label(1).Caption = "斗号"
        Label(2).Caption = "品种号"
        Label(3).Caption = "输送罐数"
        Label(4).Caption = "当前罐数"
        Label(5).Caption = "1#压力"
        Label(6).Caption = "2#压力"
        Label(7).Caption = "管道压力"
        Label(8).Caption = "时间"
        Label(9).Caption = "状态"
        
    End If
Else
    m_window.Caption = "&Window"
    m_w_liucheng.Caption = "&Close Demo Window"
    m_option.Caption = "&Option"
    m_o_zoom.Caption = "&Zoom Window"
    m_o_guandao.Caption = "&Pipeline Property"
    AddErr.Caption = "&Add Error"
    DelErr.Caption = "&Delete Error"
    If Make_Mathine <> 0 Then
        Label(0).Caption = "che"
        Label(1).Caption = "run_che"
        Label(2).Caption = "duan"
        Label(3).Caption = "temp"
        Label(4).Caption = "duan_time"
        Label(5).Caption = "total_time"
        Label(6).Caption = "power"

    Else
        Label(0).Caption = "mathine"
        Label(1).Caption = "dou"
        Label(2).Caption = "pzh"
        Label(3).Caption = "send"
        Label(4).Caption = "guan"
        Label(5).Caption = "1#power"
        Label(6).Caption = "2#power"
        Label(7).Caption = "guanpower"
        Label(8).Caption = "time"
        Label(9).Caption = "statu"
    End If
    
End If
Select Case Make_Mathine
    Case 0
        m_w_main_tk.Enabled = True
        m_w_sub_tk1.Enabled = True
        m_w_sub_tk2.Enabled = True
    Case 1
        m_w_main_tk.Enabled = False
        m_w_sub_tk1.Enabled = True
        m_w_sub_tk2.Enabled = False
    Case 2
        m_w_main_tk.Enabled = False
        m_w_sub_tk1.Enabled = False
        m_w_sub_tk2.Enabled = False
    Case Else
End Select
'-------------------------------------------------------

'MAIN_MDI.Process_Msg.Visible = False
'MAIN_MDI.ProcessBar.Visible = False
F_ShowLc.MousePointer = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Del_Win(Me.Hwnd)
    

End Sub

Private Sub m_o_errhis_Click()
    F_Error_History.Show vbModal
End Sub

Private Sub m_o_errorhelp_Click()
Dim Error_Code, Error_Prompt As String
If Not F_Error_Help.Visible Then
    ErrorGrid.row = ErrorGrid.TopRow
    ErrorGrid.col = 1
    Error_Code = Trim(ErrorGrid.Text)
    ErrorGrid.col = 2
    Error_Prompt = "(" + Error_Code + ") " + Trim(ErrorGrid.Text)
    
    ErrorData.Recordset.index = "error_idx"
    ErrorData.Recordset.Seek "=", Error_Code
    If Not ErrorData.Recordset.NoMatch Then
        F_Error_Help.Label2.Caption = Error_Prompt
        If Not IsNull(ErrorData.Recordset!China_Help) Then
            F_Error_Help.Text1.Text = ErrorData.Recordset!China_Help
        Else
            F_Error_Help.Text1.Text = ""
        End If
        F_Error_Help.Show vbModal
        Reading_ErrorHelp = True
    Else
        F_Error_Help.Label2.Caption = Error_Prompt
        F_Error_Help.Text1.Text = ""
    End If
End If
End Sub

Private Sub m_o_guandao_Click()
F_Gd_Lc.Show vbModal
End Sub

Private Sub m_o_refresh_Click()
F_ShowLc.Refresh
End Sub

Private Sub m_o_StopSpeaking_Click()
Dim i As Integer
For i = 0 To Cur_Errors.Length - 1
    Call kill_error(CLng(Cur_Errors.Code(i)))
Next i
Call sndPlaySound("", SND_ASYNC)
End Sub

Private Sub m_o_test_Click()
Dim i As Integer
Dls.Space_len = 6
Dls.Xd_len = 24
Me.Timer1.Enabled = False
Me.Timer2.Enabled = False
For i = 1 To Pic_num
    Dev_PIC(i).Enabled = True
    Dev_PIC(i).Can_Comein = True
    Dev_PIC(i).Can_Goout = True
    Dev_PIC(i).Max_Size = 100
    Dev_PIC(i).Value = 1
    Dev_PIC(i).Obj_Passing = True
Next i

'Rollback_ZLine.Length = 0
'Call Normal_Traver
'For i = 0 To Rollback_ZLine.Length - 1
'    Debug.Print Rollback_ZLine.ZIndex(i), Dev_pic(Dev_zline(Rollback_ZLine.ZIndex(i)).Start_pic_index).Name, Dev_pic(Dev_zline(Rollback_ZLine.ZIndex(i)).End_pic_index).Name
'Next i

Call CT_DFS.Test_Traver
For i = 0 To Zline_num - 1
    Call F_ShowLc.Hide_Zline(i, True)
Next i
Me.Timer1.Enabled = True
Me.Timer2.Enabled = True
End Sub

Private Sub m_o_tk_machine_Click()
F_TK_Machine.Show vbModal
End Sub

Private Sub m_o_zoom_Click()
F_Size.Show vbModal
End Sub

Private Sub m_w_error_Click()
If Not F_Error.Visible Then
    F_Error.Show vbModal
End If
F_Error.ZOrder 0
End Sub

Private Sub m_w_liucheng_Click()
Dim temp As Integer
 temp = InStr(1, MAIN_MDI.Caption, "--", 1)
 If temp <> 0 Then
    MAIN_MDI.Caption = Left(MAIN_MDI.Caption, temp - 1)
End If
Me.Enabled = False
'Timer1.Enabled = False
'Timer2.Enabled = False
Me.Hide
End Sub

Private Sub m_w_main_tk_Click()
F_Slt_Tk.Show vbModal
End Sub


'TIMER1为专门用于执行动画的定时器
Private Sub Timer1_Timer()
In_Timer1 = True
If Make_Mathine = 0 Then
    Call Run_XHD(True)                   '运行信号灯
    Call Run_FaMen(True)               '运行阀门
    Call Run_Zlines                          '运行管道
    Call Show_Status(False)         '显示状态
    Exit Sub
End If
'Call old_Timer2_Timer
Select Case Run_flag
    Case 0          '停止动画
    Case Else   '执行动画
        Call Run_XHD(True)                   '运行信号灯
        Call Run_FaMen(True)               '运行阀门
        Call Run_Zlines                          '运行管道
        Call Run_THC_YLCS(False)                  '运行碳黑秤,油料秤
        Call Run_JLCS(False)                   '运行胶料秤
    '    Call Run_ECCDS(False)               '运行二次储斗
        Call Run_LJLS(False)                  '运行炼焦炉
        Call Show_Status(False)         '显示状态
End Select
In_Timer1 = False
DoEvents
End Sub

'TIMER2为专门用于设置设备的定时器
Private Sub Timer2_Timer()
In_Timer2 = True
If Make_Mathine = 0 Then
            Call Set_FaMens_XHDS_Status             '阀门,信号灯
            'Call Set_THC_YLCS_Status                '碳黑秤,油料秤
            'Call Set_JLCS_Status                    '胶料秤
            Call Set_RCDS_Status                    '日储斗
        '    call set_ylds_status                   '油料斗
            Call Set_ECCDS_Status                   '二次储斗
            Call Luo_Traver     '遍历图库
            Call Run_Error                          '显示报警信息及声音 change by zgh 98.3
            
End If
Select Case Run_flag
    Case 0              '停止动画
    Case Else           '设置设备
 '       If Not In_Timer1 Then
            Call Set_FaMens_XHDS_Status             '阀门,信号灯
            Call Set_THC_YLCS_Status                '碳黑秤,油料秤
            Call Set_JLCS_Status                    '胶料秤
            Call Set_RCDS_Status                    '日储斗
        '    call set_ylds_status                   '油料斗
            Call Set_ECCDS_Status                   '二次储斗
            Call Luo_Traver     '遍历图库
            Call Run_Error                          '显示报警信息及声音 change by zgh 98.3
 '       End If
End Select
In_Timer2 = False
DoEvents
End Sub

'****************************************************************************
'* 函数功能：实时显示生产过程中的某些重要参数（车数、段时间、温度、能量等）     *
'****************************************************************************
Public Function Show_Status(Force_Show As Boolean)
Dim Tmp, duan_hao, Power, tempro, Flag As Integer
Static Duan_Start_Time, All_Start_Time As Long
Static Duan_Time_Count, All_Time_Count As Boolean
Static Duan_First, All_First As Boolean
Static Pre_Run_Flag As Integer
Static Fen_table As Fen_Disp_Dou
'风送部分
    If Make_Mathine = 0 Then
        Call read_fen_statu(Fen_table)
        StatusVal(0).Caption = Fen_table.Mathine
        StatusVal(1).Caption = Fen_table.Dou
        StatusVal(2).Caption = Fen_table.Kind
        StatusVal(3).Caption = Fen_table.Send_data
        StatusVal(4).Caption = Fen_table.Now_data
        StatusVal(5).Caption = Vround(Fen_table.One, 2)
        StatusVal(6).Caption = Vround(Fen_table.Two, 2)
        StatusVal(7).Caption = Vround(Fen_table.Power, 2)
        StatusVal(8).Caption = Fen_table.My_time
        StatusVal(9).Caption = Fen_table.Statu
        Exit Function
    End If
    
'Tmp = read_run(s, Flag)
'Dim s As Run_Table
'Tmp = read_run(s, Run_flag)
If Run_flag = 2 Then
    Select Case Run_Do_Table.Duan_Time
        Case 1      'start
            If Not Duan_First Then
                Duan_Start_Time = Timer
                Duan_First = True
            End If
            Duan_Time_Count = True
    Case 0     'end
            Duan_Time_Count = False
            Duan_First = False
End Select
If Duan_Time_Count Then
        StatusVal(4).Caption = Trim(CStr(CLng(Timer - Duan_Start_Time)))
Else
        StatusVal(4).Caption = "0"
End If

Select Case Run_Do_Table.All_Che_Time
    Case 1      'start
        If Not All_First Then
            All_Start_Time = Timer
            All_First = True
            End If
            All_Time_Count = True
        Case 0     'end
            All_Time_Count = False
            All_First = False
    End Select
    If All_Time_Count Then
        StatusVal(5).Caption = Trim(CStr(CLng(Timer - All_Start_Time)))
    Else
        StatusVal(5).Caption = "0"
    End If
Else ' run_flag= 1
    
If Run_Do_Table.Duan_Time > 0 Then
        StatusVal(4).Caption = Trim(CStr(Run_Do_Table.Duan_Time))
Else
        StatusVal(4).Caption = "0"
End If

    If Run_Do_Table.All_Che_Time > 0 Then
        StatusVal(5).Caption = Trim(CStr(Run_Do_Table.All_Che_Time))
    Else
        StatusVal(5).Caption = "0"
    End If

End If  'end if run_flag

StatusVal(0).Caption = Trim(CStr(Run_Do_Table.total_che))
StatusVal(1).Caption = Trim(CStr(Run_Do_Table.Now_Che))

Tmp = read_power(duan_hao, Power, tempro, Flag)
StatusVal(2).Caption = Trim(CStr(duan_hao))
StatusVal(3).Caption = Trim(CStr(tempro))
StatusVal(6).Caption = Trim(CStr(Power))

If Run_flag <> Pre_Run_Flag Then
    Select Case Run_flag
        Case 0          '停止
        Case 2
        Case Else
    End Select
    Pre_Run_Flag = Run_flag
End If
End Function


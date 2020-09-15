VERSION 4.00
Begin VB.Form lian_liao_table_input 
   BackColor       =   &H00C0C0C0&
   Caption         =   "lian_liao_table.frm"
   ClientHeight    =   4650
   ClientLeft      =   870
   ClientTop       =   1890
   ClientWidth     =   7395
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   5055
   Left            =   810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3187.755
   ScaleMode       =   0  'User
   ScaleWidth      =   7388.87
   Top             =   1545
   Width           =   7515
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Uptata"
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Record"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del Record"
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Record"
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Record"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1335
   End
   Begin VB.Data lian_liao_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      Exclusive       =   0   'False
      Height          =   615
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lain_liao_table"
      Top             =   4320
      Width           =   6495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Move Record"
      Height          =   852
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   6732
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<"
         Height          =   375
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         Height          =   375
         Left            =   6000
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label record_prompt 
         Caption         =   "Label1"
         Height          =   372
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2172
      End
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Text            =   "Txtanname"
      Top             =   6840
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      _Version        =   327680
      Tab             =   2
      TabHeight       =   529
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "lian_liao_table.frx":0000
      Tab(0).ControlCount=   12
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "pei_number_text"
      Tab(0).Control(0).Enabled=   -1  'True
      Tab(0).Control(1)=   "Text20"
      Tab(0).Control(1).Enabled=   -1  'True
      Tab(0).Control(2)=   "Text21"
      Tab(0).Control(2).Enabled=   -1  'True
      Tab(0).Control(3)=   "Text22"
      Tab(0).Control(3).Enabled=   -1  'True
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   -1  'True
      Tab(0).Control(5)=   "duan_Text"
      Tab(0).Control(5).Enabled=   -1  'True
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(11).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "lian_liao_table.frx":001C
      Tab(1).ControlCount=   16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label17"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text24"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Text23"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "Text18"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "Text17"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "Text16"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "Text15"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Text6"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Text5"
      Tab(1).Control(15).Enabled=   -1  'True
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "lian_liao_table.frx":0038
      Tab(2).ControlCount=   14
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label24"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label23"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label22"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label21"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label20"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label19"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label18"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text13"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text11"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text10"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Text9"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Text8"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Text7"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Text3"
      Tab(2).Control(13).Enabled=   0   'False
      Begin VB.TextBox pei_number_text 
         BackColor       =   &H00C0FFC0&
         DataField       =   "pei_number"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -73440
         TabIndex        =   35
         Text            =   "TxtNumber"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00C0FFC0&
         DataField       =   "password"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -70440
         TabIndex        =   34
         Text            =   "TxtPassward"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00C0FFC0&
         DataField       =   "neg"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -73440
         TabIndex        =   33
         Text            =   "TxtNeg"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H00C0FFC0&
         DataField       =   "tem"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -70440
         TabIndex        =   32
         Text            =   "TxtTem"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         DataField       =   "speed"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -73440
         TabIndex        =   31
         Text            =   "TxtSpeed"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox duan_Text 
         BackColor       =   &H00C0FFC0&
         DataField       =   "duan"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -70440
         TabIndex        =   30
         Text            =   "TxtDuan"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFC0&
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -73320
         TabIndex        =   29
         Text            =   "TxtYouname"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFC0&
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -70080
         TabIndex        =   28
         Text            =   "TxtTanName"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00C0FFC0&
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -73320
         TabIndex        =   27
         Text            =   "TxtJiaoName"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00C0FFC0&
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -69960
         TabIndex        =   26
         Text            =   "TxtXiaoname"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00C0FFC0&
         DataField       =   "dou_you_time"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -73320
         TabIndex        =   25
         Text            =   "TxtYoutime"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00C0FFC0&
         DataField       =   "dou_tan_time"
         DataSource      =   "lian_liao_Data"
         Height          =   360
         Left            =   -69960
         TabIndex        =   24
         Text            =   "TxtTantime"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H00C0FFC0&
         DataField       =   "dou_jiao_time"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -73320
         TabIndex        =   23
         Text            =   "TxtJiaotime"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text24 
         BackColor       =   &H00C0FFC0&
         DataField       =   "dou_jiao_time"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   -69960
         TabIndex        =   22
         Text            =   "TxtXiaotime"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         DataField       =   "st"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Text            =   "TxtSt"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFC0&
         DataField       =   "pre"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         Text            =   "TxtPre"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFC0&
         DataField       =   "lian_time"
         DataSource      =   "lian_liao_Data"
         Height          =   360
         Left            =   1560
         TabIndex        =   19
         Text            =   "TxtLiantime"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0FFC0&
         DataField       =   "stop_time"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   4920
         TabIndex        =   18
         Text            =   "TxtStoptime"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFC0&
         DataField       =   "upper_pre"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Text            =   "TxtUpperpre"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFC0&
         DataField       =   "ctr"
         DataSource      =   "lian_liao_Data"
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Text            =   "TxtCtr"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0FFC0&
         DataField       =   "upper_statue"
         DataSource      =   "lian_liao_Data"
         Height          =   360
         Left            =   3720
         TabIndex        =   15
         Text            =   "TxtUpperstatue"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Number:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Passward:"
         Height          =   375
         Left            =   -71760
         TabIndex        =   55
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Neg:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   54
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Tem:"
         Height          =   375
         Left            =   -71760
         TabIndex        =   53
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Speed:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   52
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FFFF&
         Caption         =   "Duan:"
         Height          =   375
         Left            =   -71760
         TabIndex        =   51
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dou_you_name:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dou_tan_name:"
         Height          =   375
         Left            =   -71640
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dou_jiao_name:"
         Height          =   372
         Left            =   -74880
         TabIndex        =   48
         Top             =   1440
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dou_xiao_name:"
         Height          =   372
         Left            =   -71640
         TabIndex        =   47
         Top             =   1440
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dou_you_time:"
         Height          =   372
         Left            =   -74880
         TabIndex        =   46
         Top             =   960
         Width           =   1572
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dou_tan_time:"
         Height          =   372
         Left            =   -71640
         TabIndex        =   45
         Top             =   960
         Width           =   1812
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dou_jiao_time:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   44
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FFFF&
         Caption         =   "Dou_xiao_time:"
         Height          =   375
         Left            =   -71640
         TabIndex        =   43
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H0080FFFF&
         Caption         =   "St:"
         Height          =   372
         Left            =   360
         TabIndex        =   42
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label19 
         BackColor       =   &H0080FFFF&
         Caption         =   "Pre:"
         Height          =   375
         Left            =   3840
         TabIndex        =   41
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackColor       =   &H0080FFFF&
         Caption         =   "Lian_time:"
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackColor       =   &H0080FFFF&
         Caption         =   "Stop_time:"
         Height          =   375
         Left            =   3840
         TabIndex        =   39
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackColor       =   &H0080FFFF&
         Caption         =   "Upper_pre:"
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackColor       =   &H0080FFFF&
         Caption         =   "Upper_statue:"
         Height          =   375
         Left            =   2040
         TabIndex        =   37
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label24 
         BackColor       =   &H0080FFFF&
         Caption         =   "Ctr:"
         Height          =   375
         Left            =   3840
         TabIndex        =   36
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   135
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   15
   End
End
Attribute VB_Name = "lian_liao_table_input"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ChangeButtons()
  cmdAdd.Visible = False
  cmdEdit.Visible = False
  cmdFind.Visible = False
  cmdDel.Visible = False
  cmdUpdate.Visible = True
  cmdCancel.Visible = True
'  txtFields(2).Visible = False
  Frame2.Visible = False
  cmdFirst.Visible = False
  cmdPrev.Visible = False
  cmdNext.Visible = False
  cmdLast.Visible = False
'  Combo2.Visible = True
End Sub

Private Sub RestoreButtons()
  cmdAdd.Visible = True
  cmdEdit.Visible = True
  cmdDel.Visible = True
  cmdFind.Visible = True
  cmdUpdate.Visible = False
  cmdCancel.Visible = False
  Frame2.Visible = True
  cmdFirst.Visible = True
  cmdPrev.Visible = True
  cmdNext.Visible = True
  cmdLast.Visible = True

End Sub

Private Sub FirstDo()
   If china_english = CHINA Then
       cmdAdd.Caption = "添加记录："
       cmdDel.Caption = "删除记录： "
       cmdEdit.Caption = "编辑记录："
       cmdFind.Caption = "查询记录："
       cmdUpdate.Caption = "更新："
       cmdCancel.Caption = "取消："
       Label1.Caption = "编号："
       Label9.Caption = "段号："
       Label3.Caption = "密码："
       Label14.Caption = "所投油次数："
       Label8.Caption = "转速："
       Label15.Caption = "所投炭次数："
       Frame2.Caption = "移动记录："
       Label10.Caption = "所投油名称："
       Label11.Caption = "所投炭名称："
       Label12.Caption = "所投胶名称："
       Label13.Caption = "所投小料名称："
       Label6.Caption = "能量"
       Label7.Caption = "温度"
       Label16.Caption = "所投胶次数："
       Label17.Caption = "所投小料次数："
       Label18.Caption = "投料上升： "
       Label19.Caption = "压力级："
       Label20.Caption = "混炼时间："
       Label21.Caption = "稳顶时间："
       Label22.Caption = "上顶栓压力："
       Label23.Caption = "上顶栓状态 ："
       Label24.Caption = "控制关系： "
    Else

   End If
   SSTab1.Tab = 0

End Sub


Public Sub cmdAdd_Click()
 Dim temp As Integer
 temp = duan_text.Text
 ChangeButtons
 lian_liao_Data.Recordset.AddNew
 lian_liao_Data.Recordset.Fields("Password").Value = password
 
 If temp > 1 Then
     duan_text.Text = temp + 1
 Else
     duan_text.Text = 1
 End If
 
 If lian_liao_Data.Recordset.RecordCount > 0 Then
   msTBookMark = lian_liao_Data.Recordset.Bookmark
 Else
   msTBookMark = ""
  End If

End Sub

Private Sub cmdCancel_Click()
lian_liao_Data.Recordset.CancelUpdate
'DTan_hei_table.Recordset.Bookmark = msTBookMark
RestoreButtons
End Sub

Private Sub cmdDel_Click()
If china_english = CHINA Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
  lian_liao_Data.Recordset.Delete
  Call cmdPrev_Click
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DTan_hei_table.Recordset.Delete
End If

End Sub

Private Sub cmdEdit_Click()
    Dim LIndex As Integer
    lian_liao_Data.RecordsetType = vbRSTypeDynaset
    lian_liao_Data.Recordset.Edit
    ChangeButtons
End Sub

Private Sub cmdFirst_Click()
    If lian_liao_Data.Recordset.RecordCount > 0 Then
      lian_liao_Data.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If lian_liao_Data.Recordset.RecordCount > 0 Then
      lian_liao_Data.Recordset.MoveLast
    End If

End Sub

Private Sub cmdNext_Click()
If lian_liao_Data.Recordset.RecordCount > 0 And Not lian_liao_Data.Recordset.BOF Then
    lian_liao_Data.Recordset.MoveNext
    If lian_liao_Data.Recordset.EOF Then
            lian_liao_Data.Recordset.MovePrevious
    End If
End If

End Sub

Private Sub cmdPrev_Click()
If lian_liao_Data.Recordset.RecordCount > 0 And Not lian_liao_Data.Recordset.BOF Then
    lian_liao_Data.Recordset.MovePrevious
    If lian_liao_Data.Recordset.BOF Then
            lian_liao_Data.Recordset.MoveNext
    End If
End If
End Sub

Private Sub cmdUpdate_Click()
  lian_liao_Data.Recordset.Update
  RestoreButtons
  If msTBookMark <> "" Then
      DTan_hei_table.Recordset.Bookmark = msTBookMark
  End If
  Exit Sub

End Sub

Private Sub Form_Load()
FirstDo
End Sub


Private Sub lian_liao_Data_Reposition()
If china_english = CHINA Then
  If Not lian_liao_Data.Recordset.EOF And Not lian_liao_Data.Recordset.BOF Then
    record_prompt.Caption = "当前记录:" & (lian_liao_Data.Recordset.AbsolutePosition + 1) & " 之 " & lian_liao_Data.Recordset.RecordCount
  End If
Else

End If

End Sub


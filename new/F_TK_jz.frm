VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form F_TK_Machine 
   ClientHeight    =   3876
   ClientLeft      =   2796
   ClientTop       =   1380
   ClientWidth     =   5604
   LinkTopic       =   "Form1"
   ScaleHeight     =   3876
   ScaleWidth      =   5604
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   2292
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5604
      _Version        =   65536
      _ExtentX        =   9885
      _ExtentY        =   4043
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
      Begin Threed.SSCommand SSCommand1 
         Height          =   492
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   78
         Caption         =   "添加记录 &A"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "F_TK_jz.frx":0000
         Height          =   2052
         Left            =   120
         OleObjectBlob   =   "F_TK_jz.frx":0010
         TabIndex        =   2
         Top             =   120
         Width           =   3852
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\ljxt\tkbmp\AllMdb.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   372
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TK_Machine"
         Top             =   1800
         Width           =   3852
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   492
         Left            =   4200
         TabIndex        =   5
         Top             =   960
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   78
         Caption         =   "删除记录 &D"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   492
         Left            =   4200
         TabIndex        =   6
         Top             =   1680
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   78
         Caption         =   "退    出 &X"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   1572
      Left            =   0
      TabIndex        =   1
      Top             =   2304
      Width           =   5604
      _Version        =   65536
      _ExtentX        =   9885
      _ExtentY        =   2773
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
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "F_TK_jz.frx":09DE
         Height          =   1332
         Left            =   120
         OleObjectBlob   =   "F_TK_jz.frx":09EE
         TabIndex        =   3
         Top             =   120
         Width           =   5412
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\ljxt\tkbmp\AllMdb.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   372
         Left            =   120
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Auth_table"
         Top             =   1080
         Width           =   4092
      End
   End
End
Attribute VB_Name = "F_TK_Machine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：设置机组与图库的对应关系

Private Sub DBGrid2_DblClick()
    DBGrid1.Columns("tk_name").Text = DBGrid2.Columns("name").Text
    DBGrid1.SetFocus
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call DBGrid2_DblClick
End If
End Sub

Private Sub Form_Activate()
F_ShowLc.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
F_ShowLc.Enabled = True
End Sub

Private Sub SSCommand1_Click()
Data1.Recordset.AddNew
End Sub

Private Sub SSCommand2_Click()
If DBGrid1.EditActive = True Then
    Speak_Error ("编辑时不能删除")
    DBGrid1.SetFocus
    Exit Sub
End If
    
If Data1.Recordset.RecordCount > 0 Then
        If Data1.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择记录")
            DBGrid1.SetFocus
            Exit Sub
      End If
        If MsgBox("真的要删除该记录吗？", 4, "记录删除") <> vbYes Then
            Exit Sub
        End If
        DBGrid1.EditActive = False
'        dbGrid1.AllowAddNew = False
        Data1.Recordset.Delete
End If
End Sub

Private Sub SSCommand3_Click()
Unload Me
End Sub


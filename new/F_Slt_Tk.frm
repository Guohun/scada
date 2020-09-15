VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form F_Slt_Tk 
   Caption         =   "选择演示系统图"
   ClientHeight    =   3240
   ClientLeft      =   2760
   ClientTop       =   2076
   ClientWidth     =   4332
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4332
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\tkbmp\AllMdb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   1200
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "TK_Machine"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1572
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "F_Slt_Tk.frx":0000
      Height          =   1812
      Left            =   240
      OleObjectBlob   =   "F_Slt_Tk.frx":0010
      TabIndex        =   2
      Top             =   600
      Width           =   3852
   End
   Begin Threed.SSCommand SSCommand1 
      Default         =   -1  'True
      Height          =   492
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   1212
      _Version        =   65536
      _ExtentX        =   2138
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "确  定 &Y"
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
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   492
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   1212
      _Version        =   65536
      _ExtentX        =   2138
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "取  消 &X"
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
   Begin VB.Label Label2 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2172
   End
   Begin VB.Label Label1 
      Caption         =   "当前演示的系统图:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
End
Attribute VB_Name = "F_Slt_Tk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：选择演示指定机组的系统图

Private Sub DBGrid1_DblClick()
Call SSCommand1_Click
End Sub

Private Sub Form_Activate()
    Data1.Recordset.index = "idx_tk_name"
    Data1.Recordset.Seek "=", Grap_Name
    If Not Data1.Recordset.NoMatch Then
        Label2.Caption = Grap_Name + " (" + str(Data1.Recordset.Fields("machine").Value) + "号机组)"
    Else
        Label2.Caption = Grap_Name
    End If
    If Data1.Recordset.RecordCount > 0 Then
        Data1.Recordset.MoveFirst
    End If
    Data1.Recordset.index = "idx_machine"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub SSCommand1_Click()
Dim Tk_Name As String
Tk_Name = DBGrid1.Columns("tk_name").Text
If (Tk_Name <> Grap_Name) And (Trim(Tk_Name) <> "") Then
    Unload F_ShowLc
    Grap_Name = Tk_Name
    F_ShowLc.Tag = Trim(DBGrid1.Columns("machine").Text)
    F_ShowLc.Show
End If
Unload Me
End Sub

Private Sub SSCommand2_Click()
Unload Me
End Sub

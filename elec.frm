VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Elec_not_table_Input 
   Caption         =   "5"
   ClientHeight    =   5292
   ClientLeft      =   -144
   ClientTop       =   1116
   ClientWidth     =   9288
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5292
   ScaleWidth      =   9288
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "elec.frx":0000
      Height          =   4692
      Left            =   6480
      OleObjectBlob   =   "elec.frx":0010
      TabIndex        =   41
      Top             =   480
      Width           =   2772
   End
   Begin VB.CommandButton exit_command 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   372
      Left            =   7560
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Width           =   1092
   End
   Begin VB.Frame Frame2 
      Caption         =   "Move Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   0
      TabIndex        =   32
      Top             =   4320
      Width           =   6372
      Begin VB.CommandButton Print_Command 
         Caption         =   "&Print"
         Height          =   375
         Left            =   2640
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   852
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&L>>"
         Height          =   375
         Left            =   5640
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&-"
         Height          =   375
         Left            =   5040
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&+"
         Height          =   375
         Left            =   4320
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&F<<"
         Height          =   375
         Left            =   3600
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.Label record_prompt 
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         ForeColor       =   &H000000FF&
         Height          =   372
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Width           =   2172
      End
   End
   Begin VB.Data DElec_not_table 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\ljxt.Mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "elec_in_table"
      Top             =   4680
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H000000C0&
      Caption         =   "&Find Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000C0&
      Caption         =   "&Del Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H000000C0&
      Caption         =   "&Add  Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Caption         =   "Frame1"
      Height          =   3612
      Left            =   0
      TabIndex        =   15
      Top             =   480
      Width           =   6372
      Begin VB.ComboBox check_list 
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   4700
         TabIndex        =   8
         Top             =   1080
         Width           =   1450
      End
      Begin VB.ComboBox run_list 
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   1500
         TabIndex        =   2
         Top             =   1080
         Width           =   1450
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "cnt"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   11
         Left            =   4920
         TabIndex        =   12
         Text            =   "txtCnt"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "use_flag"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   9
         Left            =   4700
         MaxLength       =   1
         TabIndex        =   11
         Text            =   "txtFlag"
         Top             =   3120
         Width           =   1450
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "value"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   10
         Left            =   1560
         TabIndex        =   6
         Text            =   "txtVlue"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "name"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   1500
         TabIndex        =   1
         Text            =   "txtName"
         Top             =   240
         Width           =   1450
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "check_time"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   8
         Left            =   1500
         TabIndex        =   5
         Text            =   "txtTime"
         Top             =   3120
         Width           =   1450
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "check_value"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   7
         Left            =   4700
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "txtCheck_value"
         Top             =   2400
         Width           =   1450
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "run_value"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   6
         Left            =   1500
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "txtRun_value"
         Top             =   2400
         Width           =   1450
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "check_set"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   5
         Left            =   4700
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "txtCheck_set"
         Top             =   1680
         Width           =   1450
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "check"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2880
         TabIndex        =   17
         Text            =   "txtCheck"
         Top             =   720
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "run_set"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   3
         Left            =   1500
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "txtRun_set"
         Top             =   1680
         Width           =   1450
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "run"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Text            =   "txtRun"
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "str"
         DataSource      =   "DElec_not_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   4700
         TabIndex        =   7
         Text            =   "txtStr"
         Top             =   240
         Width           =   1450
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cnt:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   11
         Left            =   3600
         TabIndex        =   31
         Top             =   3600
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vlue:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   10
         Left            =   240
         TabIndex        =   30
         Top             =   3600
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Flag:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   372
         Index           =   9
         Left            =   3300
         TabIndex        =   29
         Top             =   3120
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   0
         Left            =   100
         TabIndex        =   28
         Top             =   240
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   372
         Index           =   8
         Left            =   100
         TabIndex        =   27
         Top             =   3120
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Check_value:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   372
         Index           =   7
         Left            =   3300
         TabIndex        =   26
         Top             =   2400
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Run_Value:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   372
         Index           =   6
         Left            =   100
         TabIndex        =   23
         Top             =   2400
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Check_set:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   372
         Index           =   5
         Left            =   3300
         TabIndex        =   22
         Top             =   1680
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Check:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   372
         Index           =   4
         Left            =   3300
         TabIndex        =   21
         Top             =   1080
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Run_set:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   372
         Index           =   3
         Left            =   100
         TabIndex        =   20
         Top             =   1680
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Run:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   372
         Index           =   2
         Left            =   100
         TabIndex        =   19
         Top             =   1080
         Width           =   1400
      End
      Begin VB.Label Label 
         BackColor       =   &H80000001&
         Caption         =   "Str:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   1
         Left            =   3300
         TabIndex        =   18
         Top             =   240
         Width           =   1400
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   4680
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "error_num"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label list_name 
      BackColor       =   &H00000040&
      Caption         =   "list_name"
      ForeColor       =   &H000000FF&
      Height          =   372
      Index           =   1
      Left            =   6720
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "Elec_not_table_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msTBookMark As String
Dim gsDBName As Database
Dim gsThisSet As Recordset

Private Sub Add_list()
    Dim name_str As String * 12
    Dim Prompt_str    As String * 20
    Dim MyDB As Database, MyData As Recordset, TempRecordset As Recordset
    Dim i As Integer, j As Integer
    
'    error_list.Clear
    
    Set MyDB = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.MDB")
    Set MyData = MyDB.OpenRecordset("error_num", dbOpenSnapshot)
    Do While Not MyData.EOF
                    name_str = MyData.Fields("error_name").Value
                    If China_English = CHINA Then
                            Prompt_str = vFieldVal(MyData.Fields("china_prompt").Value)
                    Else
                            Prompt_str = vFieldVal(MyData.Fields("eng_prompt").Value)
                    End If
                    MyData.MoveNext
    Loop
    MyData.Close
    MyDB.Close
    
End Sub






Private Sub check_list_Click()
                txtFields(4).Text = check_list.List(check_list.ListIndex)
End Sub


Private Sub check_list_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtFields(4).Text = check_list.List(check_list.ListIndex)
    End If
        If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call Print_Command_Click
    End If
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Call CmdFind_Click
    End If
    If KeyCode = vbKeyF10 Then
        KeyCode = 0
        Call cmdDelete_Click
    End If
    

End Sub

Private Sub cmdAddNew_Click()

 ChangeButtons
 DElec_not_table.Recordset.AddNew
End Sub


Private Sub cmdCancel_Click()
DElec_not_table.Recordset.CancelUpdate
'DElec_not_table.Recordset.Bookmark = msTBookMark
RestoreButtons
End Sub


Private Sub cmdDelete_Click()
If DElec_not_table.Recordset.RecordCount = 0 Then
        Exit Sub
End If

If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
  DElec_not_table.Recordset.Delete
  DElec_not_table.Refresh
  Call cmdPrev_Click
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DElec_not_table.Recordset.Delete
End If
End Sub

Private Sub cmdEdit_Click()
'Dim LIndex As Integer
If DElec_not_table.Recordset.RecordCount > 1 Then
    DElec_not_table.Recordset.Edit
    ChangeButtons
End If
'LIndex = 0
'If txtFields(2).Text <> "" Then
'Do
' If (txtFields(2).Text) = combo1.List(LIndex) Then Exit Do
' LIndex = LIndex + 1
'Loop
'End If
'combo1.ListIndex = LIndex
'Call Combo1_Click
End Sub

Private Sub CmdFind_Click()
  On Error GoTo FindErr:
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
FindStart:
  For i = 0 To 8
      frmFind.lstFields.AddItem Label(i).Caption   'Mid((Label(i).Caption), 1, Len(Label(i).Caption) - 1)
  Next
  


  'reset the flags
  gbFindFailed = False
  gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOp
  frmFind.txtExpression.Text = gsFindExpr
  frmFind.optFindType(gnFindType) = True
'  mbNotFound = False
  frmFind.Show vbModal
  
  
  
 If gbFindFailed = True Then Exit Sub 'find cancelled
  '  GoTo AfterWhile
  'End If
  Select Case gsFindFiel
    Case Label(0).Caption
      sTmp = txtFields(0).DataField + space(1) + gsFindOp + space(1) + "'" + gsFindExpr + "'"
    Case Label(1).Caption
      sTmp = txtFields(1).DataField + space(1) + gsFindOp + space(1) + "'" + gsFindExpr + "'"
    Case Label(2).Caption
      sTmp = txtFields(2).DataField + space(1) + gsFindOp + space(1) + "'" + gsFindExpr + "'"
    Case Label(3).Caption
      sTmp = txtFields(3).DataField + space(1) + gsFindOp + space(1) + gsFindExpr
    Case Label(4).Caption
      sTmp = txtFields(4).DataField + space(1) + gsFindOp + space(1) + gsFindExpr
    Case Label(5).Caption
      sTmp = txtFields(5).DataField + space(1) + gsFindOp + space(1) + gsFindExpr
    Case Label(6).Caption
      sTmp = txtFields(6).DataField + space(1) + gsFindOp + space(1) + gsFindExpr
    Case Label(7).Caption
      sTmp = txtFields(7).DataField + space(1) + gsFindOp + space(1) + gsFindExpr
    Case Label(9).Caption
      sTmp = txtFields(8).DataField + space(1) + gsFindOp + space(1) + gsFindExpr
  End Select
  i = frmFind.lstFields.ListIndex
  sBookMark = DElec_not_table.Recordset.Bookmark
  'search for the record
  Select Case gnFindType
    Case 0
      DElec_not_table.Recordset.FindFirst sTmp
    Case 1
      DElec_not_table.Recordset.FindNext sTmp
    Case 2
      DElec_not_table.Recordset.FindPrevious sTmp
    Case 3
      DElec_not_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = DElec_not_table.Recordset.NoMatch
  'If mbNotFound = True Then MsgBox "not found!"
AfterWhile:

  Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    DElec_not_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    'MsgBox "Record Not Found", 48
    Call Speak_Error("没发现记录")
    DElec_not_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = DElec_not_table.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    DElec_not_table.Recordset.Bookmark = sBookMark
  End If

  mbDataChanged = False

  Exit Sub

FindErr:
'  Screen.MousePointer = vbDefault
'  If Err <> gnEOF_ERR Then
'    ShowError
    Exit Sub
''  Else
    mbNotFound = True
'    Resume Next
'  End If

End Sub




Private Sub cmdFirst_Click()
If DElec_not_table.Recordset.RecordCount > 1 Then
  DElec_not_table.Recordset.MoveFirst
End If
End Sub

Private Sub cmdLast_Click()
If DElec_not_table.Recordset.RecordCount > 1 Then
        DElec_not_table.Recordset.MoveLast
End If
End Sub

Private Sub cmdNext_Click()
If DElec_not_table.Recordset.RecordCount > 1 And Not DElec_not_table.Recordset.EOF Then
    DElec_not_table.Recordset.MoveNext
    If DElec_not_table.Recordset.EOF Then DElec_not_table.Recordset.MoveLast
End If
End Sub

Private Sub cmdPrev_Click()
If DElec_not_table.Recordset.RecordCount > 1 And Not DElec_not_table.Recordset.BOF Then
    DElec_not_table.Recordset.MovePrevious
    If DElec_not_table.Recordset.BOF Then DElec_not_table.Recordset.MoveFirst
End If
End Sub


Private Sub CmdUpdate_Click()
     For i% = 0 To 11
            If txtFields(i%).Text = "" And txtFields(i%).Visible Then
                Call Speak_Error("应输入" & Label(i%).Caption)
                txtFields(i%).SetFocus
                Exit Sub
         End If
   Next i%
  DElec_not_table.Recordset.Update
  RestoreButtons
  'DElec_not_table.Refresh
  
  
End Sub




Public Sub ChangeButtons()
  cmdAddNew.Visible = False
'  cmdEdit.Visible = False
  CmdFind.Visible = False
  cmdDelete.Visible = False
  CmdUpdate.Visible = True
  CmdCancel.Visible = True
  cmdFirst.Visible = False
  cmdPrev.Visible = False
  cmdNext.Visible = False
  cmdLast.Visible = False
End Sub

Public Sub RestoreButtons()
  cmdAddNew.Visible = True
'  cmdEdit.Visible = True
  cmdDelete.Visible = True
  CmdFind.Visible = True
  CmdUpdate.Visible = False
  CmdCancel.Visible = False
  Frame2.Visible = True
  cmdFirst.Visible = True
  cmdPrev.Visible = True
  cmdNext.Visible = True
  cmdLast.Visible = True

End Sub








Private Sub DBGrid1_DblClick()
                     txtFields(0).Text = DBGrid1.Columns("error_name").Text
                     txtFields(1).SetFocus
                     Exit Sub

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyTab Then
'            run_Grid.SetFocus
            
            Exit Sub
        End If
     If KeyCode = 13 Then
                    KeyCode = 0
                     txtFields(0).Text = DBGrid1.Columns("error_name").Text
                     txtFields(1).SetFocus
                     Exit Sub
     End If
     If (KeyCode <> 40 And KeyCode <> 38) Then KeyCode = 0


End Sub

Private Sub DElec_not_table_Reposition()
Dim i%
If China_English = CHINA Then
  If Not DElec_not_table.Recordset.EOF And Not DElec_not_table.Recordset.BOF Then
        record_prompt.Caption = "当前记录:" & (DElec_not_table.Recordset.AbsolutePosition + 1) & " 之 " & DElec_not_table.Recordset.RecordCount
        For i% = 0 To run_list.ListCount - 1
                    run_list.List(i%) = txtFields(2).Text
                    run_list.ListIndex = i%
                    Exit For
        Next i%
        For i% = 0 To run_list.ListCount - 1
                  check_list.List(i%) = txtFields(4).Text
                  check_list.ListIndex = i%
                  Exit For
        Next i%
        
  Else
      record_prompt.Caption = "无记录:"
  End If
Else

End If

End Sub


Private Sub DElec_not_table_Validate(Action As Integer, Save As Integer)
If Save = -1 Then
     For i% = 0 To 11
            If txtFields(i%).Text = "" And txtFields(i%).Visible Then
                Call Speak_Error("应输入" & Label(i%).Caption)
                txtFields(i%).SetFocus
                Save = 0
                Exit Sub
         End If
   Next i%
End If
End Sub


Private Sub error_list_DblClick()
      txtFields(0).Text = Left(error_list.List(error_list.ListIndex), 4)
End Sub


Private Sub Exit_Command_Click()
    Unload Me
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
            Width = MAIN_MDI.Width - 100 ' Set width of form.
            Height = MAIN_MDI.Height - 900 ' Set height of form.
            Left = 0
            Top = 0

Call Add_Win(Hwnd)

Call Add_list
Call FirstDo
DElec_not_table.Refresh
If DElec_not_table.Recordset.RecordCount > 0 Then
    DElec_not_table.Recordset.MoveLast
    DElec_not_table.Recordset.MoveFirst
End If

End Sub


'装入启动接点和灯
Public Sub FirstDo()
Dim i, count As Integer
run_list.Clear

For count = 0 To Elec_Output.Length         '启动
        run_list.AddItem Elec_Output.Note_Name(count)
Next count
check_list.Clear

For count = 0 To Elec_Input.Length  ' 灯
        check_list.AddItem Elec_Input.Note_Name(count)
Next count
End Sub





Private Sub Form_Paint()
If China_English = True Then
  cmdAddNew.Caption = "&A添加记录"
  Caption = "故障测量表"
  cmdDelete.Caption = "&F10删除记录"
  CmdFind.Caption = "&F查询记录"
  CmdUpdate.Caption = "&U更新"
  CmdCancel.Caption = "&C取消"
  Exit_Command.Caption = "ESC退出"
  Print_Command.Caption = "F4打印"
  Frame1.Caption = ""
  Frame2.Caption = "移动记录"
  Label(0).Caption = "错误名:"
  Label(1).Caption = "保留:"
  Label(2).Caption = "起动接点:"
  Label(3).Caption = "逻辑设定值:"
  Label(4).Caption = "检测接点:"
  Label(5).Caption = "逻辑设定值:"
  Label(6).Caption = "起动接点实际值:"
  Label(7).Caption = "检测接点实际值:"
  Label(8).Caption = "检测时间:"
  Label(9).Caption = "停检标志:"
  Label(10).Caption = "结果:"
  Label(11).Caption = "计数:"
  DBGrid1.Caption = "错误表"
  DBGrid1.Columns(0).Caption = "错误代号 "
  DBGrid1.Columns("china_prompt").Caption = "错误代号 "
  DBGrid1.Columns("eng_prompt").Visible = False
  DBGrid1.Columns("china_prompt").Visible = True
  
Else
    Caption = "Error Check"
  cmdAddNew.Caption = "&Add"
   cmdDelete.Caption = "&Delte"
  CmdFind.Caption = "&Find"
  CmdUpdate.Caption = "&Update"
  CmdCancel.Caption = "&Cancel"
  Exit_Command.Caption = "&Exit"
  Frame1.Caption = ""
  Frame2.Caption = "move"
  Label(0).Caption = "elec_disp_name:"
  Label(1).Caption = "haveing:"
  Label(2).Caption = "start:"
  Label(3).Caption = "set:"
  Label(4).Caption = "check:"
  Label(5).Caption = "check_value:"
  Label(6).Caption = "check_real:"
  Label(7).Caption = "read:"
  Label(8).Caption = "time:"
  Label(9).Caption = "flag:"
  Label(10).Caption = "real"
  Label(11).Caption = "count:"
  DBGrid1.Caption = "Error  table "
  DBGrid1.Columns(0).Caption = "error code "
  DBGrid1.Columns("eng_prompt").Caption = "prompt "
  DBGrid1.Columns("eng_prompt").Visible = True
  DBGrid1.Columns("china_prompt").Visible = False
End If

End Sub

Private Sub Form_Terminate()
        Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call Del_Win(Hwnd)
End Sub


Private Sub Print_Command_Click()
   'On Error GoTo error_doing
   Dim Y, i%, k%
   Dim temp(0 To 20) As Single
    
   Screen.MousePointer = vbHourglass
    If Printer.ScaleWidth < Me.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        Run_Grid.SetFocus
        Exit Sub
    End If
  With DElec_not_table
        Printer.Height = (.Recordset.RecordCount + 2) * txtFields(0).Height
        Printer.Width = Screen.Width
   End With
    
      temp(0) = 6
   k = 1
   
   For i = 0 To 9
                temp(k) = 0
                temp(k) = temp(k - 1) + (Printer.ScaleWidth) / 9
                k = k + 1
   Next i
   
   Printer.Print
   Printer.Print
    If China_English = CHINA Then
        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
        Printer.Print "故障错误表 " + Date$
    Else
        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
        Printer.Print "set report " + Date$
    End If
'    Printer.FontName = "宋体"
 '   Printer.FontSize = 12
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    k% = 0
    For i% = 0 To 9
               Printer.PSet (temp(k%), Printer.CurrentY)
                Printer.Print Label(i%).Caption;
                k% = k% + 1
    Next i%
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    With DElec_not_table
    .Recordset.MoveFirst
    j% = 0
    Do While Not .Recordset.EOF()
        If j% < 48 Then
            k% = 0
           For i% = 0 To 9
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print txtFields(i%).Text; 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
          Next i%
         Printer.Print String(1, " ")
         .Recordset.MoveNext
         j% = j% + 1
        Else
             Y = Printer.CurrentY + 10
            Printer.Line (0, Y)-(Printer.ScaleWidth, Y)     '  print --------
            Printer.Print
            Printer.Print "第" + CStr(Printer.Page) + "页"
            Printer.Print
            Printer.NewPage
            Printer.Print
            Printer.Print
            Y = Printer.CurrentY + 10
            Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
            Printer.Print
            j% = 0
        End If
    Loop
    .Recordset.MoveFirst
End With
'    Do While i% < 48
 '           Printer.Print
  '          i% = i% + 1
   ' Loop
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
'    Printer.Print "页", Printer.Page
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
  Screen.MousePointer = vbDefault

End Sub

Private Sub run_list_Click()
            txtFields(2).Text = run_list.List(run_list.ListIndex)
End Sub


Private Sub run_list_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
                txtFields(2).Text = run_list.List(run_list.ListIndex)
  End If
    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call Print_Command_Click
    End If
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Call CmdFind_Click
    End If
    If KeyCode = vbKeyF10 Then
        KeyCode = 0
        Call cmdDelete_Click
    End If
    
  
End Sub

Private Sub txtFields_Change(index As Integer)
    If index = 4 Then
         check_list.Text = txtFields(4).Text
    End If
    If index = 2 Then
         run_list.Text = txtFields(2).Text
    End If
    
End Sub

Private Sub txtFields_GotFocus(index As Integer)
          txtFields(index).ForeColor = Get_Focus_Fore_Color
          If index = 0 And txtFields(0).Text <> "" Then
                'Data1.Recordset.FindFirst "error_name =" & txtFields(0).Text
                
          End If
End Sub


Private Sub txtFields_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call Print_Command_Click
    End If
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Call CmdFind_Click
    End If
    If KeyCode = vbKeyF10 Then
        KeyCode = 0
        Call cmdDelete_Click
    End If
    
End Sub

Private Sub txtFields_LostFocus(index As Integer)
  txtFields(index).ForeColor = Los_Focus_Fore_Color
End Sub



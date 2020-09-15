VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Password_Form 
   Caption         =   "Pass"
   ClientHeight    =   5400
   ClientLeft      =   432
   ClientTop       =   1440
   ClientWidth     =   8256
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   8256
   Begin VB.Data comm_data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\comm.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Password"
      Top             =   5040
      Width           =   3132
   End
   Begin VB.Frame Frame2 
      Height          =   732
      Left            =   0
      TabIndex        =   5
      Top             =   4080
      Width           =   7692
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&F<<"
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&+"
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&-"
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&L>>"
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.Label record_prompt 
         Caption         =   "Label1"
         Height          =   372
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1812
      End
   End
   Begin VB.CommandButton exit_command 
      Caption         =   "&Exit"
      Height          =   372
      Left            =   6720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1092
   End
   Begin VB.CommandButton Cancel_Command 
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
      Left            =   3480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Update_Command 
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
      Left            =   1920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Delete_command 
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
      Left            =   5160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Add_Command 
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
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid Run_Grid 
      Bindings        =   "Password_Form.frx":0000
      Height          =   3375
      Left            =   0
      OleObjectBlob   =   "Password_Form.frx":0014
      TabIndex        =   11
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "Password_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FirstDo()
Dim i    As Integer

      run_Grid.Font.Size = 12
      run_Grid.HeadFont.Size = 12
        For i = 0 To 3
               run_Grid.Columns(i).Width = 2000
        Next i
        run_Grid.RowHeight = 600
        run_Grid.HeadLines = 2
        Add_Command.Enabled = True
        Update_Command.Enabled = False
      
  If china_english = CHINA Then
      run_Grid.Columns(0).Caption = "模块名"
      run_Grid.Columns(1).Caption = "用户值"
      
      run_Grid.Columns(2).Caption = "密码值"
      run_Grid.Columns(3).Caption = "登记时间"
      
    Add_Command.Caption = "&A追加"
    Update_Command.Caption = "&U修改"
    Delete_command.Caption = "&D删除"
    Cancel_Command.Caption = "&C取消"
    Exit_Command.Caption = "&E退出"
    
 Else
    
End If
End Sub






Private Sub add_Command_Click()
        Add_Command.Enabled = False
        run_Grid.AllowAddNew = True
        comm_data.Recordset.AddNew
        Update_Command.Enabled = True
        
End Sub

Private Sub cmdFirst_Click()
 comm_data.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
  comm_data.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
 If comm_data.Recordset.EOF Then
        comm_data.Recordset.MoveLast
 Else
 comm_data.Recordset.MoveNext
 End If

End Sub

Private Sub cmdPrev_Click()
 comm_data.Recordset.MovePrevious
 If comm_data.Recordset.BOF Then
        comm_data.Recordset.MoveFirst
 End If

End Sub

Private Sub comm_Data_Reposition()
If china_english = CHINA Then
  If Not comm_data.Recordset.EOF And Not comm_data.Recordset.BOF Then
    record_prompt.Caption = "当前记录:" & (comm_data.Recordset.AbsolutePosition + 1) & " 之 " & comm_data.Recordset.RecordCount
  End If
Else

End If

End Sub



Private Sub Delete_command_Click()
  If comm_data.Recordset.RecordCount > 0 Then
        comm_data.Recordset.Delete
  End If

End Sub

Private Sub exit_command_Click()
  Unload Me
End Sub

Private Sub Form_Load()
'    china_english = 1
        Call Add_Win(Hwnd)

            Width = MAIN_MDI.Width  ' Set width of form.
            Height = MAIN_MDI.Height * 3 / 4 ' Set height of form.
            Left = 0
            Top = 0
         comm_data.DatabaseName = Data_Path & "\comm.mdb"
         comm_data.Refresh
    
    
End Sub

Private Sub Print_Command_Click()
   On Error GoTo error_doing
   Dim Y
'    Printer.NewPage
    Printer.FontName = "黑体"
    Printer.FontSize = 20
    If china_english = CHINA Then
        Printer.Print String(20, " ");
        Printer.Print "设备管理报表 " + Date$
    Else
        Printer.Print String(20, " ");
        Printer.Print "set report " + Date$
    End If
    Printer.FontName = "宋体"
    Printer.FontSize = 12
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    For i% = 0 To run_Grid.VisibleCols - 1
            Printer.Print run_Grid.Columns(i%).Caption + String(6, " ");
    Next i%
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    For i% = 0 To run_Grid.VisibleRows - 1
'        run_Grid.SelBookmarks.Add run_Grid.RowBookmark(i%)
        If i% < 48 Then
            run_Grid.row = i%
            For j% = 0 To run_Grid.VisibleCols - 1
                    Printer.Print run_Grid.Columns(j%).Value + String(6, " ");
            Next j%
            Printer.Print String(1, " ")
        Else
            Printer.Print "页", Printer.Page
            Printer.NewPage
            i% = 0
        End If
    Next i%
    Do While i% < 48
            Printer.Print
            i% = i% + 1
    Loop
    Printer.Print "页", Printer.Page
    Printer.EndDoc
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
End Sub


Private Sub Form_Paint()
Call FirstDo
End Sub

Private Sub Form_Terminate()
Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Del_Win(Hwnd)

End Sub


Private Sub run_Grid_Error(ByVal DataError As Integer, Response As Integer)
            Response = 0

End Sub

Private Sub UPDATE_Command_Click()
        Add_Command.Enabled = True
'        comm_Data.Recordset.Update
        run_Grid.AllowAddNew = False
        
        Update_Command.Enabled = False
End Sub


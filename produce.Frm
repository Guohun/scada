VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form produce 
   Caption         =   "Produce"
   ClientHeight    =   6792
   ClientLeft      =   -576
   ClientTop       =   1320
   ClientWidth     =   9492
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6792
   ScaleWidth      =   9492
   Begin Threed.SSCommand Simulate_Command 
      Height          =   732
      Left            =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "Simulate"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
      Picture         =   "produce.frx":0000
   End
   Begin MSDBGrid.DBGrid run_Grid 
      Bindings        =   "produce.frx":0452
      Height          =   3012
      Left            =   120
      OleObjectBlob   =   "produce.frx":0465
      TabIndex        =   1
      Top             =   960
      Width           =   8052
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "produce.frx":1CEC
      Height          =   1740
      Left            =   0
      OleObjectBlob   =   "produce.frx":1CFC
      TabIndex        =   2
      Top             =   4080
      Width           =   8172
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   8280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "SSCommand1"
      ForeColor       =   255
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
      Picture         =   "produce.frx":357E
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   4320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "Find"
      ForeColor       =   12582912
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
      Picture         =   "produce.frx":39D0
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "Insert"
      ForeColor       =   12582912
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
      Picture         =   "produce.frx":3E22
   End
   Begin Threed.SSCommand Del_Command 
      Height          =   732
      Left            =   7080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "SSCommand1"
      ForeColor       =   12582912
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
      MouseIcon       =   "produce.frx":4274
      Picture         =   "produce.frx":46C6
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "&Print"
      ForeColor       =   12582912
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
      MouseIcon       =   "produce.frx":4B18
      Picture         =   "produce.frx":4F6A
   End
   Begin Threed.SSCommand Change_Command 
      Height          =   852
      Left            =   8280
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1092
      _Version        =   65536
      _ExtentX        =   1926
      _ExtentY        =   1503
      _StockProps     =   78
      Caption         =   "SSCommand1"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Font3D          =   4
      Picture         =   "produce.frx":53BC
   End
   Begin Threed.SSCommand cmdRefresh 
      Height          =   732
      Left            =   5880
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "SSCommand1"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
      Picture         =   "produce.frx":580E
   End
   Begin VB.Data run_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Run_Table"
      Top             =   3600
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\LJXT\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "pei_fan_table"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label list_name 
      BackColor       =   &H00000000&
      Caption         =   "list_name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2652
      Left            =   8400
      TabIndex        =   0
      Top             =   960
      Width           =   492
   End
End
Attribute VB_Name = "produce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim make_only_key As Integer '当前make表中的only_key记录
Private Sub Add_list()
    Dim number_str    As String * 18, name_str As String * 15
    Dim MyDB As Database, MyData As Recordset
    Dim i As Integer, j As Integer

    Exit Sub
    Set MyDB = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.MDB")
    Set MyData = MyDB.OpenRecordset("pei_fan_table", dbOpenSnapshot)
    
    If MyData.RecordCount > 0 Then
         MyData.MoveFirst
         Do While Not MyData.EOF
                         number_str = vFieldVal(MyData.Fields(1).Value)
                         name_str = vFieldVal(MyData.Fields(2).Value)
  '                       name_str = namestr & MyData.Fields(2).Value
                        pei_list.AddItem Trim(number_str) + String$(8 - Len(Trim(number_str)), "_") + "_" + name_str
                        MyData.MoveNext
        Loop
        pei_list.ListIndex = 0
    End If
    MyData.Close
    MyDB.Close
End Sub





Private Sub Change_Command_Click()
On Error Resume Next
    Dim temp_number   As String
    Dim temp_name   As String
    Dim temp_ban     As Integer
    Dim temp_che    As Integer
    Dim temp_flag     As Integer
    Dim Mathine%
    Dim book    As Variant
    
    If Run_flag = 0 Then
            Call Speak_Error("处于停止状态 ")
            Run_Grid.SetFocus
            Exit Sub
    End If
    run_Data.UpdateRecord
'If run_Grid.EditActive = True Then
 '           Speak_Error ("编辑后要确认才能在线修改")
  '          run_Grid.SetFocus
   '         Exit Sub
 'End If
             
                         book = Run_Grid.Bookmark
                         run_Data.Recordset.FindFirst "index=" & Make_Row
                        For i = 0 To Run_Grid.Columns.count - 1
                                If Run_Grid.Columns(i).Visible = True And Run_Grid.Columns(i).Caption = "" Then
                                        Call Speak_Error("第一条记录" & Run_Grid.Columns(i).Caption & "没有输入")
                                        Exit Sub
                                End If
                        Next i
                    If Run_Grid.Columns("_number") = "" Or IsEmpty(Run_Grid.Columns("_number")) Then
                        Call Speak_Error("第" & Make_Row & "记录的" & " 配方编号没有输入")
                        Exit Sub
                    End If
                    If Run_Grid.Columns("ban") = "" Or IsNull(Run_Grid.Columns("ban")) Then
                        Call Speak_Error("第" & Make_Row & "记录的" & " 班号没有输入")
                        Exit Sub
                    End If
                    If Run_Grid.Columns("che") = "" Or IsNull(Run_Grid.Columns("che")) Then
                        Call Speak_Error("第" & Make_Row & "记录的" & " 车数没有输入")
                        Exit Sub
                    End If
                   
                         
                         
                         temp_number = vFieldVal(run_Data.Recordset.Fields("_number").Value)
                         temp_name = vFieldVal(run_Data.Recordset.Fields("name").Value)
                         temp_ban = Val(run_Data.Recordset.Fields("ban").Value)
                         temp_che = Val(run_Data.Recordset.Fields("che").Value)
                         Mathine = Val(run_Data.Recordset.Fields("mathine").Value)
                     If (temp_che <= Run_Do_Table.Now_Che + 1) And (Run_Grid.Columns("index").Value = Make_Row) Then
                          Speak_Error ("修改的车数必须大于目前的生产车数:" & Run_Do_Table.Now_Che)
                            Run_Grid.SetFocus
                            Exit Sub
                    End If
                         
                          Call can_read(temp_number, temp_name, temp_ban, temp_che, 48 Or Mathine) 'change
                          Run_Grid.Bookmark = book
             
            
End Sub

Private Sub CmdFind_Click()
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
FindStart:
  
   
   frmFind.lstFields.AddItem Run_Grid.Columns("_number").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("name").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("m_day").Caption

  'reset the flags
  gbFindFailed = False
  gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOp
  frmFind.txtExpression.Text = gsFindExpr
  frmFind.optFindType(gnFindType) = True
  frmFind.Show vbModal

  
  
 If gbFindFailed = True Then Exit Sub 'find cancelled
  '  GoTo AfterWhile
  'End If
  sTmp = "[" + Run_Grid.Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
  
  sBookMark = run_Data.Recordset.Bookmark
  'search for the record
  Select Case gnFindType
    Case 0
      run_Data.Recordset.FindFirst sTmp
    Case 1
      run_Data.Recordset.FindNext sTmp
    Case 2
      run_Data.Recordset.FindPrevious sTmp
    Case 3
      run_Data.Recordset.FindLast sTmp
  End Select
  mbNotFound = run_Data.Recordset.NoMatch
  If mbNotFound = True Then MsgBox "not found!"
AfterWhile:
  Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    run_Data.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    run_Data.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = run_Data.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    run_Data.Recordset.Bookmark = sBookMark
  End If

  mbDataChanged = False
  Run_Grid.SetFocus
  
  Exit Sub

FindErr:
'  Screen.MousePointer = vbDefault
'  If Err <> gnEOF_ERR Then
'    ShowError
    error_name.SetFocus
    
    Exit Sub
''  Else
    mbNotFound = True
'    Resume Next
'  End If

End Sub

Private Sub cmdRefresh_Click()
    If Make_Mathine <> 0 Then
        run_Data.UpdateRecord
        Prev_run_flag = False
        cmdRefresh.Enabled = False
        Simulate_Command.Enabled = False
        MAIN_MDI!Timer2.Enabled = True
    Else
        Call Speak_Error("主机不可运行生产命令，只可运行风送系统")
        Run_Grid.SetFocus
    End If
End Sub


Private Sub DBGrid1_DblClick()
                On Error Resume Next
                     Run_Grid.Columns("_NUMBER").Text = DBGrid1.Columns("_number").Text
                     Run_Grid.Columns("NAME").Text = DBGrid1.Columns("name").Text
                     Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                     Run_Grid.SetFocus
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.col = 1
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error Resume Next
     Select Case KeyCode
        Case vbKeyTab
                Exit Sub
        Case 13
                     Run_Grid.Columns("_NUMBER").Text = DBGrid1.Columns("_number").Text
                     Run_Grid.Columns("NAME").Text = DBGrid1.Columns("name").Text
                     Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                     KeyCode = 0
                     DBGrid1.SetFocus
                     Exit Sub
        Case vbKeyF5
            If Run_flag = 0 Then
                Call simulate_command_Click
            End If
            KeyCode = 0

        Case vbKeyF7
            If Run_flag = 0 Then
                Call cmdRefresh_Click
            End If
            KeyCode = 0
            
        Case vbKeyF2
            If Run_flag <> 0 Then
                Call Change_Command_Click
            End If
            KeyCode = 0
            Run_Grid.SetFocus
     End Select
     If (KeyCode <> 40 And KeyCode <> 38) Then KeyCode = 0

End Sub

Private Sub DBGrid1_LostFocus()
Run_Grid.col = 0
End Sub

Private Sub Del_Command_Click()
Dim i%
    
On Error GoTo ErrorHandler  ' Enable error-handling routine.
        
If Run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能删除")
            Run_Grid.SetFocus
            Exit Sub
 End If
    
    If Run_Grid.Columns("index").Value <= Make_Row Then
                         'Speak_Error ("须选择配方")
                                Speak_Error (Run_Grid.Columns("_number") + "配方已生产")
                                
                                run_Data.Recordset.FindFirst "index>" & Make_Row
                                Run_Grid.Refresh
                                Run_Grid.SetFocus
                                Exit Sub
    End If
    
  If run_Data.Recordset.RecordCount > 0 Then
        If run_Data.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择配方")
            Run_Grid.SetFocus
            Exit Sub
      End If
        If MsgBox("真的要删除该记录吗？", 4, "记录删除") <> vbYes Then
            Exit Sub
        End If
        Run_Grid.EditActive = False
        Run_Grid.AllowAddNew = False
        i% = Run_Grid.Columns(0).Value
        run_Data.Recordset.Delete
        Dim strsqlrestore   As String
        strsqlrestore = "update run_table  set  index=index-1   where index>=" & i%
        
        With run_Data
                .Database.Execute strsqlrestore, dbFailOnError
                .Refresh
                .Recordset.FindFirst "index=" & i%
            If .Recordset.NoMatch Then
                    If .Recordset.RecordCount > 0 Then
                            .Recordset.MoveLast
                    End If
            End If
            
        End With
        Run_Grid.AllowUpdate = True
        Run_Grid.AllowAddNew = True
        Run_Grid.SetFocus
        SendKeys "{HOME}"
  End If

  Exit Sub
ErrorHandler:
    Speak_Error ("出错:")
    Run_Grid.SetFocus
End Sub



Private Sub Exit_Command_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
        Dim SQL As String
        Call Add_Win(Hwnd)
        FirstDo
            Width = MAIN_MDI.Width  ' Set width of form.
            
            Height = MAIN_MDI.Height * 3 / 4 ' Set height of form.
            Left = 0
            Top = 0
                    
        Me.WindowState = 2
        run_Data.DatabaseName = Data_Path & "\ljxt.mdb"
        run_Data.RecordSource = "select * from   run_table order by  index"
        run_Data.Refresh
        If run_Data.Recordset.RecordCount = 0 Then
         run_Data.Recordset.AddNew
          run_Data.Refresh
          Run_Grid.Columns(0).Value = 1             'index
          Run_Grid.Columns("m_time").Text = time
          Run_Grid.Columns("mathine").Value = 1
          Run_Grid.Refresh
          Simulate_Command.Enabled = False
          cmdRefresh.Enabled = False
  End If
  If Run_flag <> 0 Then
          Simulate_Command.Enabled = False
          cmdRefresh.Enabled = False
          Change_Command.Enabled = True
  End If
  Call Add_list
End Sub


Public Sub FirstDo()
Dim i    As Integer
      Run_Grid.Columns(5).Visible = False
      Run_Grid.Columns(7).Visible = False
      
      Run_Grid.Columns(0).Locked = True         'index
      Run_Grid.Columns(2).Locked = True         'name
      Run_Grid.Columns(6).Locked = True         'my_time
'      Run_Grid.Columns(10).Locked = True         'mathine
      Run_Grid.Columns(8).Visible = False
      Run_Grid.Columns(9).Visible = False
      Run_Grid.Columns(10).Visible = True
        For i = 0 To Run_Grid.VisibleCols - 1
               Run_Grid.Columns(i).Width = 1000
        Next i
               Run_Grid.Columns("name").Width = 2000
        Run_Grid.Font.Size = 10
        Run_Grid.HeadFont.Size = 12
        Run_Grid.RowHeight = 600
        Run_Grid.HeadLines = 2
        
        
        DBGrid1.Columns(0).Visible = False
        DBGrid1.Columns(3).Visible = False
        DBGrid1.Columns(4).Visible = False
        DBGrid1.Columns(5).Visible = False
        DBGrid1.Columns(6).Visible = False
        DBGrid1.Columns(7).Visible = False
        DBGrid1.Columns(8).Visible = False
        DBGrid1.Columns(9).Visible = False
        DBGrid1.Columns("_number").Visible = True
        DBGrid1.Columns("name").Visible = True
        
        
        DBGrid1.Font.Size = 10
        DBGrid1.HeadFont.Size = 10
        DBGrid1.RowHeight = 300
        DBGrid1.HeadLines = 2

End Sub

Private Sub Form_Paint()
  If China_English = CHINA Then
     Caption = "生产命令输入与管理"
      Run_Grid.Columns(0).Caption = "目录序号"
      Run_Grid.Columns(1).Caption = "配方编号"
      Run_Grid.Columns(2).Caption = "配方名称"
      Run_Grid.Columns(3).Caption = "班    号"
      Run_Grid.Columns(4).Caption = "车    数"
      
      Run_Grid.Columns(5).Caption = "日期"
      Run_Grid.Columns(6).Caption = "时间"
      Run_Grid.Columns(7).Caption = "标    志"
      Run_Grid.Columns(8).Caption = "当前车数"
      Run_Grid.Columns(9).Caption = "当前段数"
      Run_Grid.Columns(10).Caption = "机组号"
      
        DBGrid1.Columns("_number").Caption = "配方编号"
        DBGrid1.Columns("name").Caption = "配方名称"
        DBGrid1.Columns("mathine").Caption = "机组"
     
     Change_Command.Caption = "F2在线修改"
     
     Del_Command.Caption = "F10删除"
    CmdFind.Caption = "F6查找"
    Print_Command.Caption = "F4打印"
    Insert_Command.Caption = "F3插入"
    Simulate_Command.Caption = "F5模拟"
    cmdRefresh.Caption = "F7运行"
    Exit_Command.Caption = "Esc退出"
    
    list_name.Caption = "TAB键切换窗口"
    
 Else
    Caption = "Procdure Input and Mange"
    Run_Grid.Caption = "coman"
      Run_Grid.Columns(0).Caption = "index"
      Run_Grid.Columns(1).Caption = "code"
      Run_Grid.Columns(2).Caption = "name"
      Run_Grid.Columns(3).Caption = "ban "
      Run_Grid.Columns(4).Caption = "che"
      
      Run_Grid.Columns(5).Caption = "date"
      Run_Grid.Columns(6).Caption = "time"
      Run_Grid.Columns(7).Caption = "flag"
      Run_Grid.Columns(10).Caption = "mathine"
      
        DBGrid1.Columns("_number").Caption = "A"
        DBGrid1.Columns("name").Caption = "B"
        DBGrid1.Columns("mathine").Caption = "C"
        
        
    CmdFind.Caption = "F6 search"
    Print_Command.Caption = "F4 print"
    Insert_Command.Caption = "F3 insert"
 
       cmdRefresh.Caption = "F6 Send"
        Exit_Command.Caption = "Close"
     Del_Command.Caption = "F10 Del"
     Simulate_Command.Caption = "F5 simulate"
     Change_Command.Caption = "F2 edit"
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim answer As Integer
Select Case Run_flag
    Case 0
        Cancel = 0
    Case 1
       Cancel = 0
    Case 2
        answer = MsgBox(" 在模拟  强行退出吗？退出时要等待", 48 + vbYesNo, "提示")
       If answer = vbYes Then
            If Simulate_Hwnd <> 0 Then
                Call TerminateProcess(Simulate_Hwnd, 0)
            End If
            Simulate_Hwnd = 0
            Call Sleep(1000)
            Call write_run(Run_Do_Table, 0)
            Call Sleep(500)
            Simulate_Command.Enabled = True
            cmdRefresh.Enabled = True
            Run_flag = 0

           Unload Me
           Unload F_ShowLc
           Cancel = 0
  Else
            Cancel = 1
  End If
End Select
  

End Sub

Private Sub Form_Resize()
On Error Resume Next
  'If WindowState <> 1 Then   'not minimized
   ' Run_Grid.Height = Height * 3 / 4 - 800 - Run_Grid.Top
    '    DBGrid1.Top = Run_Grid.Height + Run_Grid.Top + 200
 ' End If
 'Run_Grid.SetFocus
End Sub









Private Sub Form_Terminate()
Call Del_Win(Hwnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Del_Win(Hwnd)

End Sub



Private Sub Insert_Command_Click()
 Dim temp%
 Dim strsqlrestore   As String
 
 On Error GoTo ErrorHandler:
 If Run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能插入")
            Run_Grid.SetFocus
            Exit Sub
 End If

  If run_Data.Recordset.RecordCount > 0 Then
        If run_Data.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
            Run_Grid.SetFocus
            Exit Sub
      End If
    temp% = Val(vFieldVal(Run_Grid.Columns(0).Value))
    If Run_Grid.Columns("index").Value <= Make_Row Then
                         'Speak_Error ("须选择配方")
                                Speak_Error (Run_Grid.Columns("_number") + "配方已生产")
                                
                                run_Data.Recordset.FindFirst "index>" & Make_Row
                                Run_Grid.Refresh
                                Run_Grid.SetFocus
                                Exit Sub
    End If
      

'        run_Grid.AllowAddNew = True
        i% = temp%
        'Dim strsqlrestore   As String
        strsqlrestore = "update run_table  set  index=index+1   where index>=" & temp%
        run_Data.Database.Execute strsqlrestore, dbFailOnError
        
        'run_Data.Recordset.FindFirst "index=" & i%
        'Do While Not run_Data.Recordset.NoMatch
         '       run_Data.Recordset.Edit
          '      run_Data.Recordset.Fields("index").value = i% + 1
           '     run_Data.Recordset.Update
            '    i% = i% + 1
             '   run_Data.Recordset.FindNext "index=" & i%
        'Loop
        
        run_Data.Recordset.AddNew
        run_Data.Recordset.Fields("index").Value = temp%
        run_Data.Recordset.Fields("m_time").Value = time
        run_Data.Recordset.Fields("mathine").Value = 1
        run_Data.Recordset.Update
        run_Data.Refresh
        
        run_Data.Recordset.FindFirst "index=" & temp%
        Run_Grid.Refresh
        Run_Grid.AllowAddNew = False
  End If
        Run_Grid.SetFocus
        
  Exit Sub
ErrorHandler:
    Speak_Error ("须选择配方")
    Run_Grid.SetFocus

End Sub

Private Sub Print_Command_Click()
   On Error GoTo error_doing

   Dim Y  As Long
   Dim i%, k%
   Dim temp(0 To 20)     As Single
   Dim book   As Variant
   book = Run_Grid.Bookmark
   
   Screen.MousePointer = vbHourglass
   
       If Printer.ScaleWidth < Run_Grid.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        Run_Grid.SetFocus
        Exit Sub
    End If
  With run_Data
        Printer.Height = (.Recordset.RecordCount + 2) * Run_Grid.RowHeight
        Printer.Width = Screen.Width
   End With
    
      temp(0) = 6
   k = 1
   With Run_Grid
   For i = 0 To .Columns.count - 1
         'temp(i) = i * Printer.ScaleWidth / run_Grid.VisibleCols
         If .Columns(i).Visible Then
                temp(k) = 0
                temp(k) = temp(k - 1) + .Columns(i).Width * (Printer.ScaleWidth) / .Width
                k = k + 1
        End If
   Next i
   End With

   
 
    Printer.Print
    Printer.Print
    'Printer.FontName = "Courier"
    'Printer.FontSize = 18
    If China_English = CHINA Then
        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
        Printer.Print Run_Grid.Caption + space(4) + Date$
    Else
        Printer.Print String(20, " ");
        k% = 0
        For i% = 0 To Run_Grid.Columns.count - 1
            If Run_Grid.Columns(i%).Visible = True Then
                    Printer.Print Trim(Run_Grid.Columns(i%).Caption); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    
            End If
        Next i%
        Printer.Print String(4, " ")
    End If
    'Printer.FontName = "Courier"
    'Printer.FontSize = 12
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
        k% = 0
        For i% = 0 To Run_Grid.Columns.count - 1
            If Run_Grid.Columns(i%).Visible = True Then
                   'Printer.Print SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    Printer.Print Trim(Run_Grid.Columns(i%).Caption); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    
            End If
        Next i%
    
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    run_Data.Recordset.MoveFirst
    j% = 0
    Do While Not run_Data.Recordset.EOF()
        If j% < 48 Then
            k% = 0
           For i% = 0 To Run_Grid.Columns.count - 1
              If Run_Grid.Columns(i%).Visible = True Then
'                    Printer.Print SetRightString(Run_Grid.Columns(i%).value, temp(K%));
                    Printer.Print Trim(Run_Grid.Columns(i%).Value); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    
              End If
          Next i%
         Printer.Print String(1, " ")
         run_Data.Recordset.MoveNext
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
        
   ' Do While j% < 48
    '        Printer.Print
     '       j% = j% + 1
    'Loop
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    'Printer.Print Space(10) + "总数:" & CStr(run_Data.Recordset.RecordCount)
    'Printer.Print "第" + CStr(Printer.Page) + "页"
    'Printer.NewPage
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    Run_Grid.Bookmark = book
    Exit Sub
error_doing:
    'MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
      Run_Grid.Bookmark = book
End Sub

Private Sub run_Data_Error(DataErr As Integer, Response As Integer)
    Select Case DataError
        Case 3024
                Call write_error(0, "没发现运行表")
    End Select
    Response = 0 'continue
End Sub









Private Sub run_Grid_AfterColEdit(ByVal ColIndex As Integer)
        On Error Resume Next
            Select Case ColIndex
                Case 3
                    If Run_Grid.Columns(3).Value = "" Then
                        Run_Grid.Columns(3).Value = 1
                    End If
                    If Run_Grid.Columns(3).Value > 10 Then
                        Speak_Error ("必须输入班别1~9")
                        Run_Grid.Columns(3).Value = 1
                        Run_Grid.SetFocus
                    End If
                Case 4
                    If Run_Grid.Columns(4).Text = "" Then
                        Run_Grid.Columns(4).Value = 1
                    End If
                    If Run_Grid.Columns(4).Value < 1 Or Run_Grid.Columns(4).Value > 100 Then
                        Speak_Error ("必须输入车1~100 之间")
                        Run_Grid.Columns(4).Value = 1
                            Run_Grid.SetFocus
                    End If
                    
                Case 10
                    If Run_Grid.Columns("mathine").Value = "" Then
                                            Run_Grid.Columns(10).Value = 1
                   End If
            End Select

            
End Sub




Private Sub Run_Grid_AfterInsert()
    If Run_flag = 0 Then
              Simulate_Command.Enabled = True
              cmdRefresh.Enabled = True
    End If
End Sub

Private Sub Run_Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
      On Error GoTo Exit_Do
        If Run_flag <> 0 And Run_Grid.Columns("index").Value <= Make_Row Then
            If Run_Grid.Columns(ColIndex).DataField <> "che" Then
                        Cancel = True
                        Exit Sub
             Else
                    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入车1~9")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
            End If
         End If
            Select Case ColIndex
                Case 3      'ban
                
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入班别1~9")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                Case 4       'ban che
            
                    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入车1~9")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                Case 10     'mathine
                    If KeyAscii > Asc("2") Or KeyAscii < Asc("0") And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入机号0~2")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
            End Select
            Exit Sub
Exit_Do:
                        Speak_Error ("Run_Grid_BeforeColEditerror")
                        Run_Grid.SetFocus
End Sub

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
      On Error GoTo Exit_Do
         
           If ColIndex = 1 Then
                If Run_Grid.Columns(1).Value = "" Then
                        Speak_Error ("必须输入配方编号")
                        Cancel = True
                        Run_Grid.SetFocus
                        Exit Sub
                 End If
                If Run_Grid.Columns(1).Value <> OldValue Then
                    Data1.Recordset.index = "number_idx"
                    Data1.Recordset.Seek "=", Run_Grid.Columns(1).Value
                    If Data1.Recordset.NoMatch Then
                                Speak_Error ("无此配方编号")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                   End If
                    
                    Call DBGrid1_DblClick
                    Run_Grid.SetFocus
                End If
           End If
            If ColIndex = 3 Or ColIndex = 4 Then
                    If Not Check_Float(Run_Grid.Columns(ColIndex).Value) Then
                                Speak_Error ("数据非法")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                    End If
            End If
            If ColIndex = 10 Then
                    Data1.Recordset.index = "mathine_number"
                    Data1.Recordset.Seek "=", Run_Grid.Columns(1).Value, Run_Grid.Columns(10).Value
                    If Data1.Recordset.NoMatch Then
                                Speak_Error ("无此配方编号")
                                Cancel = True
                                Run_Grid.SetFocus
                                Data1.Recordset.index = "number_idx"
                                Exit Sub
                    End If
                    
                    
            End If
            Exit Sub
Exit_Do:
                        Speak_Error ("Run_Grid_BeforeColUpdateerror")
                        Run_Grid.SetFocus
    
End Sub











Private Sub run_Grid_BeforeUpdate(Cancel As Integer)
On Error GoTo Exit_Do
        If Run_flag <> 0 And Run_Grid.Columns("index").Value <= Make_Row Then
                    If (Run_Grid.Columns("che").Value <= (Run_Do_Table.Now_Che + 1)) And (Run_Grid.Columns("index").Value = Make_Row) Then
                        Speak_Error ("修改的车数必须大于目前的当前生产车数:" & Run_Do_Table.Now_Che)
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
         End If
            Exit Sub
Exit_Do:
                        Speak_Error ("Run_Grid_BeforeColUpdateerror")
                        Run_Grid.SetFocus
             
End Sub

Private Sub Run_Grid_Error(ByVal DataError As Integer, Response As Integer)
'        Speak_Error (DataError)
        Response = 0
End Sub

Private Sub run_Grid_GotFocus()
        If Run_Grid.col = 0 Then
                Run_Grid.col = 1
        End If

End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    Dim i, Flag As Integer
    Dim old_index As Integer
    Select Case KeyCode
        Case vbKeyDown
            Flag = 0
            For i = 0 To Run_Grid.Columns.count - 1
              If Run_Grid.Columns(i).Value = "" And Run_Grid.Columns(i).Visible Then
                    Flag = 1
                    Exit For
              End If
            Next i
            
            
            If Not IsNull(Run_Grid.Columns("index").Value) And Flag = 0 Then
                    old_index = Val(Run_Grid.Columns(index).Value)
                    If old_index >= run_Data.Recordset.RecordCount Then
                        run_Data.UpdateRecord
                        old_index = old_index + 1
                      run_Data.Recordset.AddNew
                      run_Data.Recordset.Fields("index").Value = old_index
                      run_Data.Recordset.Fields("m_time").Value = time
                      run_Data.Recordset.Fields("mathine").Value = 1
                      run_Data.Recordset.Update
                    Else
                        If Val(Run_Grid.Columns("index").Value) <= Make_Row Then
                                Speak_Error (Run_Grid.Columns("_number") + "配方已生产")
                                Run_Grid.SetFocus
                         End If
                    End If
            End If
            Run_Grid.AllowAddNew = False
            Exit Sub
errorHandle:
        Speak_Error ("keydown error ")
        Run_Grid.SetFocus

        Case vbKeyUp
                        Run_Grid.AllowAddNew = False
                        If Val(Run_Grid.Columns("index").Value) <= Make_Row Then
                                Speak_Error (Run_Grid.Columns("_number") + "配方已生产")
                                KeyCode = 0
                                Run_Grid.SetFocus
                         End If
        Case vbKeyF3
            Call Insert_Command_Click
            Run_Grid.SetFocus
            KeyCode = 0
        Case vbKeyF4
            Call Print_Command_Click
            Run_Grid.SetFocus
            KeyCode = 0
        Case vbKeyF5
            If Run_flag = 0 Then
                Call simulate_command_Click
            End If
            'Run_Grid.SetFocus
            KeyCode = 0
        Case vbKeyF6
            Call CmdFind_Click
            Run_Grid.SetFocus
            KeyCode = 0
        Case vbKeyF7
            If Run_flag = 0 Then
                Call cmdRefresh_Click
            End If
            KeyCode = 0
            
        Case vbKeyF10
            Call Del_Command_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF2
            If Run_flag <> 0 Then
                Call Change_Command_Click
            End If
            KeyCode = 0
            Run_Grid.SetFocus
            
    End Select
    
    
        
    
    
End Sub


Private Sub run_Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
        On Error Resume Next
              If Run_Grid.VisibleRows < 1 Then
                    Exit Sub
              End If
            If Run_Grid.col = 0 Then
                        Run_Grid.col = 1
            End If
            If Run_Grid.col >= 5 Then
                    If (LastCol) > 6 Then
                        Run_Grid.col = 4
                    Else
                        Run_Grid.col = 10
                    End If
            End If
    
End Sub

'模拟按钮命令
'只模拟第一条记录
Private Sub simulate_command_Click()
          On Error Resume Next
          Dim Mathine%
            Dim Z As Long
    If Make_Mathine = 0 Then
        If Fen_Run_flag <> 0 And Fen_Run_flag <> 10 Then
            Call Speak_Error("在运行， 不可模拟风送系统")
            Me.SetFocus
            Exit Sub
            
        End If
        If Run_Hwnd <> 0 Then
                            Call TerminateProcess(Run_Hwnd, 0)
                            Run_Hwnd = 0
        End If
        
        Call write_Fs_command(1)
        
        cmdRun.Enabled = False
        Timer1.Enabled = True
        Me.Hide
        Screen.MousePointer = vbHourglass
        F_ShowLc.Show
        SendKeys "%4", True  ' Send UpKey to close Calculator.
        SendKeys "1", True  ' Send UpKey to close Calculator.

        Simulate_Hwnd = ExecProgram(MAIN_MDI.Hwnd, Data_Path & "\simulate.exe")
            
        Screen.MousePointer = vbDefault
        Do While True
            If Fen_Run_flag = 0 And Fen_Run_flag = 1 Then
            GoTo simulate_exit
            End If
            DoEvents
        Loop
        
        Exit Sub
    End If
        
            
If Run_flag = 1 Then
                             Call Speak_Error("正在运行")
                             GoTo simulate_exit
End If
If Run_Grid.VisibleRows < 1 Then
                             Call Speak_Error("没输入配方")
                             GoTo simulate_exit
End If
            run_Data.UpdateRecord
             Screen.MousePointer = vbHourglass
              If Run_flag = 0 Then
                        run_Data.Recordset.FindFirst "index=1"
                        For i = 0 To Run_Grid.Columns.count - 1
                                If Run_Grid.Columns(i).Visible = True And Run_Grid.Columns(i).Caption = "" Then
                                        Call Speak_Error("第一条记录" & Run_Grid.Columns(i).Caption & "没有输入")
                                        GoTo simulate_exit
                                End If
                        Next i
                    If Run_Grid.Columns("_number") = "" Or IsEmpty(Run_Grid.Columns("_number")) Then
                        Call Speak_Error("第" & Make_Row & "记录的" & " 配方编号没有输入")
                        GoTo simulate_exit
                    End If
                    If Run_Grid.Columns("ban") = "" Or IsNull(Run_Grid.Columns("ban")) Then
                        Call Speak_Error("第" & Make_Row & "记录的" & " 班号没有输入")
                        GoTo simulate_exit
                    End If
                    If Run_Grid.Columns("che") = "" Or IsNull(Run_Grid.Columns("che")) Then
                        Call Speak_Error("第" & Make_Row & "记录的" & " 车数没有输入")
                        GoTo simulate_exit
                    End If
                   
                         
                         temp_number = vFieldVal(Run_Grid.Columns("_number").Value)
                         temp_name = vFieldVal(Run_Grid.Columns("name").Value)
                         temp_ban = Val(Run_Grid.Columns("ban").Value)
                         temp_che = Val(Run_Grid.Columns("che").Value)
                         Mathine = Val(Run_Grid.Columns("mathine").Value)
                         If Run_Hwnd <> 0 Then
                            Call TerminateProcess(Run_Hwnd, 0)
                            Run_Hwnd = 0
                         End If
                         If can_read(temp_number, temp_name, temp_ban, temp_che, 32 Or Mathine) = 0 Then
                            Call Speak_Error("没设置ODBC 或ljxt.mdb数据库有错误")
                            GoTo simulate_exit
                         End If
                         Run_flag = 2
                        
            End If
             Me.Hide
             F_ShowLc.Show
             SendKeys "%4", True  ' Send UpKey to close Calculator.
             SendKeys "1", True  ' Send UpKey to close Calculator.
             Simulate_Command.Enabled = False
             cmdRefresh.Enabled = False
             Call Sleep(1000)
             Change_Command.Enabled = True
             Simulate_Hwnd = ExecProgram(MAIN_MDI.Hwnd, Data_Path & "\simulate.exe")
             
             Screen.MousePointer = vbDefault
            
    '等待模拟 结束
    Do While True
        If Run_flag = 0 Then
            If Simulate_Hwnd = 0 Then GoTo simulate_exit
            Call Sleep(1000)
            If MsgBox("模拟结束，继续模拟吗？", vbYesNo + vbDefault) = vbYes Then
                         run_Data.Recordset.FindFirst "index=1"
                        For i = 0 To Run_Grid.Columns.count - 1
                                If Run_Grid.Columns(i).Visible = True And Run_Grid.Columns(i).Caption = "" Then
                                        Call Speak_Error("第一条记录" & Run_Grid.Columns(i).Caption & "没有输入")
                                        GoTo simulate_exit
                                End If
                        Next i
                         temp_number = vFieldVal(Run_Grid.Columns("_number").Value)
                         temp_name = vFieldVal(Run_Grid.Columns("name").Value)
                         temp_ban = Val(Run_Grid.Columns("ban").Value)
                         temp_che = Val(Run_Grid.Columns("che").Value)
                         Mathine = Val(Run_Grid.Columns("mathine").Value)
                          Call can_read(temp_number, temp_name, temp_ban, temp_che, 32 Or Mathine)
                          Run_flag = 2
            Else
                GoTo simulate_exit
            End If
        End If
        DoEvents
    Loop
simulate_exit:
            Call TerminateProcess(Simulate_Hwnd, 0)
            Simulate_Hwnd = 0
            Call Sleep(1000)
            If Run_Hwnd = 0 Then
                Run_Hwnd = ExecProgram(MAIN_MDI.Hwnd, Data_Path & "\ljm.exe")
            End If
            Run_Grid.Refresh
            Simulate_Command.Enabled = True
            cmdRefresh.Enabled = True
            Change_Command.Enabled = False
            Screen.MousePointer = vbDefault
            'run_Grid.SetFocus
End Sub



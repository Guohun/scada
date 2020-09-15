VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form error_Input 
   Caption         =   "error_input"
   ClientHeight    =   6000
   ClientLeft      =   108
   ClientTop       =   816
   ClientWidth     =   9096
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   9096
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   5052
      Left            =   7440
      TabIndex        =   0
      Top             =   840
      Width           =   1572
      Begin Threed.SSCommand China_Command 
         Height          =   730
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Tag             =   "wave"
         Top             =   360
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   1926
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
         Font3D          =   2
         Picture         =   "error.frx":0000
      End
      Begin Threed.SSCommand Eng_Command 
         DragIcon        =   "error.frx":045A
         Height          =   730
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Tag             =   "wave"
         Top             =   1440
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   1926
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
         Font3D          =   2
         Picture         =   "error.frx":08A4
      End
      Begin Threed.SSCommand China_Command 
         Height          =   730
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   1926
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
         Font3D          =   2
         Picture         =   "error.frx":0CFE
      End
      Begin Threed.SSCommand Eng_Command 
         Height          =   730
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   3840
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   1926
         _ExtentY        =   1288
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
         Font3D          =   2
         Picture         =   "error.frx":1150
      End
   End
   Begin VB.Data DError_num_table 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "error_num"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "error.frx":15A2
      Height          =   5052
      Left            =   120
      OleObjectBlob   =   "error.frx":15BD
      TabIndex        =   1
      Top             =   840
      Width           =   7212
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   6840
      TabIndex        =   2
      Top             =   0
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2346
      _ExtentY        =   1288
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
      Font3D          =   3
      Picture         =   "error.frx":3587
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   3840
      TabIndex        =   3
      Top             =   0
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2346
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
      Font3D          =   3
      Picture         =   "error.frx":39D9
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2346
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
      Font3D          =   3
      Picture         =   "error.frx":3E2B
   End
   Begin Threed.SSCommand cmdDel 
      Height          =   732
      Left            =   5400
      TabIndex        =   5
      Top             =   0
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2350
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
      Font3D          =   3
      MouseIcon       =   "error.frx":427D
      Picture         =   "error.frx":46CF
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2350
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
      Font3D          =   3
      MouseIcon       =   "error.frx":4B21
      Picture         =   "error.frx":4F73
   End
End
Attribute VB_Name = "error_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
 'Dim Error_Db   As Database, Error_Rs As Recordset



Private Sub FirstDo()
End Sub



Private Sub china_command_Click(index As Integer)
    If index = 0 Then
       If DBGrid1.Columns(0).Value <> "" Then
            frmWave.Tag = Data_Path & "\wave\china\" & Trim(DBGrid1.Columns(0).Value) & ".wav"
            frmWave.Show
            frmWave.Caption = DBGrid1.Columns("china_prompt").Value & "错误录音"
       Else
            Call Speak_Error("没错误代号")
      End If
    Else
         frmErrHelp.Show
         frmErrHelp!ErrorHelp.Text = vFieldVal(DError_num_table.Recordset.Fields("china_help").Value)
         frmErrHelp.Tag = "china_help"
    End If
End Sub




Private Sub cmdDel_Click()
Dim book   As Variant
Dim temp%
Dim temp_row As Integer
Dim i%
Dim strsqlrestore As String



If China_English = CHINA Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
   
         Screen.MousePointer = vbHourglass
        
       
        
        temp% = DError_num_table.Recordset.Fields("error_name").Value
'        temp_row = DError_num_table.Recordset.AbsolutePosition
        
        DError_num_table.Recordset.Delete
        
  '                    i% = 0
'                      With DError_num_table
 '                   .Refresh
'                  Do While Not .Recordset.EOF()
 '                           .Recordset.Edit
  '                          .Recordset.Fields("error_name").Value = i% + 1
   '                         .Recordset.Update
    '                        i% = i% + 1
     '                       .Recordset.MoveNext
      '              Loop
       '       End With
        strsqlrestore = "update error_num set  error_name=error_name-1 where error_name>" & temp%
        DError_num_table.Database.Execute strsqlrestore, dbFailOnError
        DError_num_table.Refresh
            DError_num_table.Recordset.FindFirst "error_name=" & temp
            If DError_num_table.Recordset.NoMatch Then
                    If Not DError_num_table.Recordset.EOF() Then
                            DError_num_table.Recordset.MoveLast
                    End If
           End If
         
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DError_num_table.Recordset.Delete
End If
Screen.MousePointer = vbDefault

End Sub


Private Sub CmdFind_Click()
On Error GoTo FindErr:
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
FindStart:
  
   
   frmFind.lstFields.AddItem DBGrid1.Columns(0).Caption
   frmFind.lstFields.AddItem DBGrid1.Columns("china_prompt").Caption
   frmFind.lstFields.AddItem DBGrid1.Columns("eng_prompt").Caption

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
  Select Case gsFindFiel
    Case DBGrid1.Columns(0).Caption
      sTmp = "error_name" + space(1) + gsFindOp + space(1) + gsFindExpr
    Case DBGrid1.Columns("china_prompt").Caption
      sTmp = "china_prompt" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
    Case DBGrid1.Columns("eng_prompt").Caption
      sTmp = "eng_prompt" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
  End Select
  i = frmFind.lstFields.ListIndex
  sBookMark = DError_num_table.Recordset.Bookmark
  'search for the record
  Select Case gnFindType
    Case 0
      DError_num_table.Recordset.FindFirst sTmp
    Case 1
      DError_num_table.Recordset.FindNext sTmp
    Case 2
      DError_num_table.Recordset.FindPrevious sTmp
    Case 3
      DError_num_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = DError_num_table.Recordset.NoMatch
  If mbNotFound = True Then MsgBox "not found!"
AfterWhile:
  Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    DError_num_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    DError_num_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = DError_num_table.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    DError_num_table.Recordset.Bookmark = sBookMark
  End If

  mbDataChanged = False
  DBGrid1.SetFocus
  
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


Private Sub CmdUpdate_Click()
        DError_num_table.Recordset.Update
        DError_num_table.Refresh
        strsqlrestore = "update error_num set  error_name=error_name+1 where error_name>= " & temp%
        DError_num_table.Database.Execute strsqlrestore, dbFailOnError
End Sub


Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim str As String
Dim i%
Dim rst  As Recordset
            Select Case DBGrid1.Columns(ColIndex).DataField
                Case "error_name"
                       If Not Check_Float(DBGrid1.Columns(ColIndex).Value) Then
                            Call Speak_Error("必须输入正确" + DBGrid1.Columns(ColIndex).Caption)
                            Cancel = True
                            DBGrid1.SetFocus
                            Exit Sub
                     End If
                     If DBGrid1.Columns(ColIndex).Value = OldValue Then
                                Exit Sub
                     End If
                     With DError_num_table
                        .Database.Execute "select error_name  into  [temp_Error]  from Error_num " & _
                         " ", dbFailOnError
                                 Set rst = .Database.OpenRecordset("select  * from [temp_Error] where error_name=" & DBGrid1.Columns(ColIndex).Value)
                                 If Not rst.EOF() Then
                                        Call Speak_Error(DBGrid1.Columns(ColIndex).Caption + " 机组口号重")
                                        rst.Close
                                        .Database.Execute "DROP TABLE [temp_Error];"
                                        DBGrid1.SetFocus
                                        Cancel = True
                                        
                                        Exit Sub
                               End If
                     'Next i
                     rst.Close
                     .Database.Execute "DROP TABLE [temp_Error];"
                    End With
                Case "china_prompt", "eng_prompt"
                       If DBGrid1.Columns(ColIndex).Text = "" Then
                            Call Speak_Error("必须输入" + DBGrid1.Columns(ColIndex).Caption)
                            Cancel = True
                      Else
                            str = Trim(DBGrid1.Columns(ColIndex).Text)
                            Debug.Print Len(str), str
                            
                            If Len(str) > 39 Then
                                Call Speak_Error(DBGrid1.Columns(ColIndex).Caption + "长度为0~40")
                                Cancel = True
                            Else
                                DBGrid1.Columns(ColIndex).Text = Trim(str)
                            End If
                            
                     End If
            End Select
End Sub

Private Sub DBGrid1_Error(ByVal DataError As Integer, Response As Integer)
            Debug.Print DataError
            Debug.Print DBGrid1.ErrorText
              Speak_Error ("不能输入相同错误号")
              DBGrid1.SetFocus
            Response = 0
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, Flag As Integer
    Select Case KeyCode
        Case vbKeyDown, vbKeyPageDown
            Flag = 0
            For i = 0 To DBGrid1.Columns.count - 1
              If DBGrid1.Columns(i).Value = "" And DBGrid1.Columns(i).Visible = True Then
                    Flag = 1
                    Exit For
              End If
            Next i
            If Flag = 0 Then
                DBGrid1.AllowAddNew = True
            Else
                DBGrid1.AllowAddNew = False
            End If
         Case vbKeyEscape
                Unload Me
        Case vbKeyF10
            Call cmdDel_Click
            KeyCode = 0
            DBGrid1.SetFocus
        Case vbKeyF3
            Call Insert_Command_Click
            DBGrid1.SetFocus
        Case vbKeyF6
            Call CmdFind_Click
            KeyCode = 0
            DBGrid1.SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            KeyCode = 0
            DBGrid1.SetFocus
    End Select
    
        
        

End Sub


Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
              Dim old_index As Integer
              On Error GoTo errorHandle
              
              If DBGrid1.VisibleRows < 1 Then
                    Exit Sub
              End If
        'If DBGrid1.col = 0 And LastCol > 0 Then
                'DBGrid1.col = LastCol
        'End If
              
            If DBGrid1.Columns("error_name").Value = "" Then
                    old_index = DError_num_table.Recordset.RecordCount + 1
                    DBGrid1.Columns("error_name").Value = old_index
            End If
            Exit Sub
errorHandle:
        Speak_Error ("须选择配方")

End Sub

Private Sub DError_num_table_Reposition()
'If China_english = CHINA Then
  If Not DError_num_table.Recordset.EOF And Not DError_num_table.Recordset.BOF Then
'        record_prompt.Caption = "当前记录:" & (DError_num_table.Recordset.AbsolutePosition + 1) & " 之 " & DError_num_table.Recordset.RecordCount
Else
 '       record_prompt.Caption = "没有记录:"
  End If
'Else

'End If

End Sub

Private Sub DError_num_table_Validate(Action As Integer, Save As Integer)
   If Save = -1 Then
        If DBGrid1.Columns(0).Text = "" Then
                         Speak_Error (DBGrid1.Columns(0).Caption & "不能为空")
                         Save = 0
                         'Action = 0
                        DBGrid1.SetFocus
                        Exit Sub
        End If
        If DBGrid1.Columns("china_prompt").Value = "" Then
                        Call Speak_Error(DBGrid1.Columns("china_prompt").Caption & "不能为空")
                         Save = 0
                         'Action = 0
                        DBGrid1.SetFocus
                        Exit Sub
        End If
        If DBGrid1.Columns("eng_prompt").Value = "" Then
                        Call Speak_Error(DBGrid1.Columns("eng_prompt").Caption & "不能为空")
                         Save = 0
                         'Action = 0
                        DBGrid1.SetFocus
                        Exit Sub
        End If

End If

End Sub


Private Sub eng_command_Click(index As Integer)
 If index = 0 Then
    If DBGrid1.Columns(0).Value <> "" Then
            frmWave.Tag = Data_Path & "\wave\english\" & Trim(DBGrid1.Columns(0).Value) & ".wav"
            frmWave.Show
            frmWave.Caption = DBGrid1.Columns("eng_prompt").Value & "Record  Sound"
     Else
            Call Speak_Error("没错误代号")
    End If
Else
        frmErrHelp.Show
        frmErrHelp!ErrorHelp.Text = vFieldVal(DError_num_table.Recordset.Fields("english_help").Value)
        frmErrHelp.Tag = "eng_help"
End If
End Sub


Private Sub Exit_Command_Click()
        Unload Me
End Sub

Private Sub Form_Activate()
        Me.WindowState = 2
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
On Error Resume Next
            Width = MAIN_MDI.Width  ' Set width of form.
            Height = MAIN_MDI.Height - 200 ' Set height of form.
            Left = 0
            Top = 0
            
            Call Add_Win(Hwnd)

  '     Set Error_Db = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.mdb")
 '      Set Error_Rs = Error_Db.OpenRecordset("error_num")
'       Set DError_num_table.Recordset = Error_Rs
        

        DError_num_table.RecordSource = "select * from error_num order by  error_name "
        DError_num_table.Refresh
        DBGrid1.Font.Size = 10
        DBGrid1.HeadFont.Size = 12
        DBGrid1.RowHeight = 400
        DBGrid1.HeadLines = 2
        
'        DError_num_table.Recordset.Index = "error_idx"
        DError_num_table.Database.Execute "DROP TABLE [temp_Error];"
End Sub

Private Sub Form_Paint()
 If China_English = CHINA Then
     Caption = " 报警表录入目录 "
     cmdDel.Caption = "F10删除记录"
     Insert_Command.Caption = "F3插入记录"
     CmdFind.Caption = "F6查询记录"
     Print_Command.Caption = "F4打印"
     China_Command(1).Caption = " 中文错误帮助"
     Eng_Command(1).Caption = "英文错误帮助"
     
     China_Command(0).Caption = " 中文声音处理"
     Eng_Command(0).Caption = " 英文声音处理"
     Exit_Command.Caption = "ESC退出"
     
    DBGrid1.Caption = "错误表"
    DBGrid1.Columns(0).Caption = "错误代号 "
    DBGrid1.Columns("china_prompt").Caption = "中文错误提示 "
    DBGrid1.Columns("eng_prompt").Caption = "英文错误提示 "
    DBGrid1.Columns("china_prompt").Visible = True
Else
     Caption = " index "
     cmdDel.Caption = "F10 delete"
     Insert_Command.Caption = "F3 insert"
     CmdFind.Caption = "F6 search"
     Print_Command.Caption = "F4 print"
     China_Command(1).Caption = " china help"
     Eng_Command(1).Caption = "eng help"
     
     China_Command(0).Caption = " china"
     Eng_Command(0).Caption = " eng"
     Exit_Command.Caption = "ESC "
     
    DBGrid1.Caption = "error table"
    DBGrid1.Columns(0).Caption = "code "
    DBGrid1.Columns("china_prompt").Caption = "china "
    DBGrid1.Columns("eng_prompt").Caption = "eng "
    DBGrid1.Columns("china_prompt").Visible = True


End If

End Sub


Private Sub Form_Terminate()
    Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call Del_Win(Hwnd)
End Sub




Private Sub Insert_Command_Click()
Dim temp%
'Dim temp_row, temp_first As Variant
Dim i%
Dim strsqlrestore As String
      Screen.MousePointer = vbHourglass

      temp% = DBGrid1.Columns(0).Value
      i% = temp%
        'temp_row = DBGrid1.row
        'temp_first = DBGrid1.FirstRow
'        DError_num_table.Database.index.Close
 '       DError_num_table.Database.Execute "drop  INDEX Error_idx " _
  '      & "ON error_num "
        
        strsqlrestore = "update error_num set  error_name=error_name+1   where error_name>=" & temp%
        DError_num_table.Database.Execute strsqlrestore, dbFailOnError
        
   '     DError_num_table.Database.Execute "CREATE UNIQUE INDEX Error_idx " _
    '    & "ON error_num (error_name   desc ) " _
     '   & "WITH primary;"
        
        DError_num_table.Recordset.AddNew
        DError_num_table.Recordset.Fields(0).Value = temp%
        DError_num_table.Recordset.Update
    
        'ProgressBar1.Visible = False
        DError_num_table.Refresh
          DError_num_table.Recordset.FindFirst "error_name=" & temp%
         'DBGrid1.FirstRow = temp_first
         'DBGrid1.row = temp_row
         DBGrid1.col = 10

        Screen.MousePointer = vbDefault
        
        'DError_num_table.Recordset.Fields("error_name").Value = i%
        'DError_num_table.Refresh
End Sub

Private Sub Print_Command_Click()
   'On Error GoTo error_doing
   Dim Y  As Long
   Screen.MousePointer = vbHourglass
   Dim temp(0 To 20)     As Integer
   
   
   If Printer.ScaleWidth < DBGrid1.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        DBGrid1.SetFocus
        Exit Sub
    End If
    
  With DError_num_table
      Debug.Print (.Recordset.RecordCount + 2) * DBGrid1.RowHeight

        If (.Recordset.RecordCount + 2) * DBGrid1.RowHeight > 40000 Then
                Printer.Height = 40000
        Else
                Printer.Height = (.Recordset.RecordCount + 2) * DBGrid1.RowHeight
        End If
        Printer.Width = Screen.Width
   End With
    
   temp(0) = 6
   temp(1) = Printer.ScaleWidth / 4
   temp(2) = Printer.ScaleWidth * 2 / 4
   temp(3) = Printer.ScaleWidth * 3 / 4
   
'    Printer.NewPage
 
    Printer.Print
    Printer.Print
    'Printer.FontName = "Courier"
    'Printer.FontSize = 20
    'Printer.FontBold = True
      Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
     Printer.Print DBGrid1.Caption + space(4) + Date$
     'Printer.FontBold = False
    'Printer.FontName = "Courier"
    'Printer.FontSize = 12
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
        k% = 0
        For i% = 0 To DBGrid1.Columns.count - 1
            If DBGrid1.Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(DBGrid1.Columns(i%).Caption); 'SetRightString(DBGrid1.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                    
            End If
        Next i%
    
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    DError_num_table.Recordset.MoveFirst
    j% = 0
    Do While Not DError_num_table.Recordset.EOF()
        If j% < 48 Then
            k% = 0
           For i% = 0 To DBGrid1.Columns.count - 1
              If DBGrid1.Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(DBGrid1.Columns(i%).Value); 'SetRightString(DBGrid1.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                    
              End If
          Next i%
           Printer.Print String(1, " ")
           DError_num_table.Recordset.MoveNext
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
            Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
            If China_English = CHINA Then
                    Printer.Print DBGrid1.Caption + "(续表)" + space(4) + Date$
            Else
                    Printer.Print DBGrid1.Caption + "(Next Table)" + space(4) + Date$
            End If
            Y = Printer.CurrentY + 10
            Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
            Printer.Print
            j% = 0
        End If
    Loop
        
    'Do While j% < 48
     '       Printer.Print
     '       j% = j% + 1
    'Loop
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
'    Printer.Print space(10) + "总数:" & CStr(DError_num_table.Recordset.RecordCount)
 '   Printer.Print "第" + CStr(Printer.Page) + "页"
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
End Sub

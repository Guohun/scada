VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFen_Pei_table 
   Caption         =   "风送系统配方"
   ClientHeight    =   5388
   ClientLeft      =   1104
   ClientTop       =   336
   ClientWidth     =   9096
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5388
   ScaleWidth      =   9096
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9000
      Top             =   1800
   End
   Begin MSDBGrid.DBGrid Run_Grid 
      Bindings        =   "frmFen_Pei_table.frx":0000
      Height          =   2412
      Left            =   240
      OleObjectBlob   =   "frmFen_Pei_table.frx":015A
      TabIndex        =   0
      Top             =   960
      Width           =   8748
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   8040
      TabIndex        =   2
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
      Picture         =   "frmFen_Pei_table.frx":184D
   End
   Begin Threed.SSCommand CmdSort 
      Height          =   732
      Left            =   2520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "Sort"
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
      Picture         =   "frmFen_Pei_table.frx":1C9F
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   120
      TabIndex        =   4
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
      Picture         =   "frmFen_Pei_table.frx":20F1
   End
   Begin Threed.SSCommand Del_Command 
      Height          =   732
      Left            =   6840
      TabIndex        =   5
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
      MouseIcon       =   "frmFen_Pei_table.frx":2543
      Picture         =   "frmFen_Pei_table.frx":2995
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   1320
      TabIndex        =   1
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
      MouseIcon       =   "frmFen_Pei_table.frx":2DE7
      Picture         =   "frmFen_Pei_table.frx":3239
   End
   Begin VB.Data datPrimaryRS 
      Align           =   2  'Align Bottom
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Fen_Pei_table"
      Top             =   5064
      Visible         =   0   'False
      Width           =   9096
   End
   Begin Threed.SSCommand cmdRun 
      Height          =   732
      Left            =   4200
      TabIndex        =   7
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
      Picture         =   "frmFen_Pei_table.frx":368B
   End
   Begin VB.Data Cai_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cai_liao_table"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2292
   End
   Begin Threed.SSCommand Change_Command 
      Height          =   732
      Left            =   5400
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1092
      _Version        =   65536
      _ExtentX        =   1926
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
      Enabled         =   0   'False
      Font3D          =   4
      Picture         =   "frmFen_Pei_table.frx":3ADD
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmFen_Pei_table.frx":3F2F
      Height          =   1500
      Left            =   240
      OleObjectBlob   =   "frmFen_Pei_table.frx":3F42
      TabIndex        =   6
      Top             =   3480
      Width           =   8508
   End
End
Attribute VB_Name = "frmFen_Pei_table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Change_Command_Click()
        Call write_Fs_command(1)
        'Call Order_Send
        'Run_Grid.Refresh
        Run_Grid.SetFocus
        
        'Change_Command.Enabled = False
End Sub

Private Sub CmdFind_Click()
On Error GoTo FindErr
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
  With datPrimaryRS.Recordset
  
 For i = 1 To Run_Grid.Columns.count - 1
            frmFind.lstFields.AddItem Run_Grid.Columns(i).Caption
  Next
  
FindStart:
  gbFindFailed = False
  gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOp
  frmFind.txtExpression.Text = gsFindExpr
  frmFind.optFindType(gnFindType) = True

  frmFind.Show vbModal

  
  
 If gbFindFailed = True Then Exit Sub 'find cancelled
 If Run_Grid.Columns(gsFindFiel).DataField <> "name" And Run_Grid.Columns(gsFindFiel).DataField <> "cai_number" Then
        sTmp = "[" + Run_Grid.Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + Trim(gsFindExpr)
Else
        sTmp = "[" + Run_Grid.Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + "'" + Trim(gsFindExpr) + "'"
End If
  
   sBookMark = .Bookmark

  Select Case gnFindType
    Case 0
      .FindFirst sTmp
    Case 1
      .FindNext sTmp
    Case 2
      .FindPrevious sTmp
    Case 3
      .Recordset.FindLast sTmp
  End Select
  mbNotFound = .Recordset.NoMatch
  If mbNotFound = True Then MsgBox "not found!"
AfterWhile:

  Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    .Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    .Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = .Bookmark  'save the new position
  
    .Bookmark = sBookMark
  End If

  mbDataChanged = False
    End With
  Exit Sub
FindErr:
    mbNotFound = True
    Exit Sub


End Sub


Private Sub CmdUpdate_Click()
  datPrimaryRS.UpdateRecord
  datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.LastModified
End Sub


Private Sub cmdRun_Click()
    If Make_Mathine <> 0 Then
        Call Speak_Error("非主机，不可单独运行风送")
        
        Me.SetFocus
        Exit Sub
    End If
        Call Order_Send
        Run_Grid.Refresh
        Call write_Fs_command(1)
        cmdRun.Enabled = False
        Change_Command.Enabled = True
        Timer1.Enabled = True

        Me.Hide
        F_ShowLc.Show
        SendKeys "%4", True  ' Send UpKey to close Calculator.
        SendKeys "1", True  ' Send UpKey to close Calculator.
        
End Sub

Private Sub CmdSort_Click()
        
        'datPrimaryRS.Database.Execute "drop  INDEX Sort_Idx "
        Call Fen_Song_Api
        Call Order_Send
        
        datPrimaryRS.RecordSource = "select * from fen_pei_table order by sort_idx "
        datPrimaryRS.Refresh
        Run_Grid.Refresh
End Sub

Private Sub datPrimaryRS_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'Throw away the error
End Sub

Private Sub datPrimaryRS_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  datPrimaryRS.Caption = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
End Sub

Private Sub datPrimaryRS_Validate(Action As Integer, Save As Integer)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
      Screen.MousePointer = vbDefault
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub DBGrid1_DblClick()
                    Run_Grid.Columns("cai_number").Text = DBGrid1.Columns("cai_number").Text
                    Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                    Run_Grid.Columns("name").Text = DBGrid1.Columns("name").Text
                    Run_Grid.Columns("dou").Text = DBGrid1.Columns("dou").Text
                    Run_Grid.SetFocus
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
           If KeyCode = vbKeyTab Then
            'run_Grid.SetFocus
                Exit Sub
        End If
     If KeyCode = 13 Then
                    Run_Grid.Columns("cai_number").Text = DBGrid1.Columns("cai_number").Text
                    Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                    Run_Grid.Columns("name").Text = DBGrid1.Columns("name").Text
                    Run_Grid.Columns("dou").Text = DBGrid1.Columns("dou").Text
                    Run_Grid.SetFocus
                   Run_Grid.col = 5
  End If
 If (KeyCode <> 40 And KeyCode <> 38) Then KeyCode = 0

End Sub

Private Sub Del_Command_Click()
Dim i%
    
On Error GoTo ErrorHandler  ' Enable error-handling routine.
        
If Run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能删除")
            Run_Grid.SetFocus
            Exit Sub
 End If
    
With datPrimaryRS
  If .Recordset.RecordCount > 0 Then
        If .Recordset.AbsolutePosition < 0 Then
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
        .Recordset.Delete
        Dim strsqlrestore   As String
        strsqlrestore = "update Fen_pei_table  set  Sort_idx=Sort_idx-1   where Sort_idx>=" & i%
        .Database.Execute strsqlrestore, dbFailOnError
        
        Run_Grid.AllowUpdate = True
        Run_Grid.AllowAddNew = True
        Run_Grid.SetFocus
        .Recordset.FindFirst "Sort_idx=" & i%
        SendKeys "{HOME}"
  End If
End With
  Exit Sub
ErrorHandler:
    Speak_Error ("出错:")
    Run_Grid.SetFocus

End Sub

Private Sub Exit_Command_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
 Width = MAIN_MDI.Width  ' Set width of form.
 Height = MAIN_MDI.Height - 600 ' Set height of form.
 Left = 0
 Top = 0
 Call Add_Win(Hwnd)
        For i = 0 To Run_Grid.Columns.count - 1
            If Run_Grid.Columns(i).Visible = True Then
                        Run_Grid.Columns(i).Width = 800
            End If
        Next i
        Run_Grid.Columns("name").Width = 1400
        Run_Grid.Font.Size = 10
        Run_Grid.HeadFont.Size = 12
        Run_Grid.RowHeight = 400
        Run_Grid.HeadLines = 2

    datPrimaryRS.RecordSource = "select * from fen_pei_table order by sort_idx"
    datPrimaryRS.Refresh
    
 Cai_Data.RecordSource = "select * from cai_liao_table  where  wna=1 and " & "   cai_number <> '  '   " & " order by dou"
 Cai_Data.Refresh
 If Fen_Run_flag <> 0 Then
     cmdRun.Enabled = False
     Change_Command.Enabled = True
 End If
End Sub

Private Sub Form_Paint()
If China_English = CHINA Then
        Run_Grid.Caption = "风送系统"
      Run_Grid.Columns(0).Caption = "序号"
      Run_Grid.Columns("cai_number").Caption = "材料编号"
      Run_Grid.Columns("name").Caption = "材料名称"
      Run_Grid.Columns(3).Caption = "品种号"
    
      Run_Grid.Columns("send_data").Caption = "输送罐数"
      Run_Grid.Columns("dou").Caption = "斗号"
      Run_Grid.Columns("mathine").Caption = "机组号"
      Run_Grid.Columns(8).Caption = "备用"
      
     Del_Command.Caption = "F10删除"
    CmdSort.Caption = "F6排序"
    Print_Command.Caption = "F4打印"
    Insert_Command.Caption = "F3插入"
        Change_Command.Caption = "F7在线修改"
        cmdRun.Caption = "F7运行"
    Exit_Command.Caption = "Esc退出"

           DBGrid1.Columns("cai_number").Caption = "材料编号"
           DBGrid1.Columns("name").Caption = "材料名称"
           DBGrid1.Columns("wna").Caption = "秤名"
           DBGrid1.Columns("dou").Caption = "斗名"
           DBGrid1.Columns("jiao").Caption = "胶状"
           DBGrid1.Columns("mathine").Caption = "机组"

    
 Else
        Run_Grid.Caption = "System"
        Run_Grid.Columns(0).Caption = "Sort_Idx"
        Run_Grid.Columns("mathine").Caption = "mathine"
        Run_Grid.Columns("cai_number").Caption = "cai_code"
        Run_Grid.Columns(3).Caption = "kind"
        Run_Grid.Columns("dou").Caption = "dou"
        Run_Grid.Columns("name").Caption = "name"
        Run_Grid.Columns(7).Caption = "send date"
        Run_Grid.Columns(8).Caption = "temp"
        Run_Grid.Columns("send_data").Caption = "dd"
        
    cmdRun.Caption = "F7Run"
    Exit_Command.Caption = "ESC Close"
     Del_Command.Caption = "F10 Del"
    CmdSort.Caption = "F6 Sort"
    Print_Command.Caption = "F4 "
    Insert_Command.Caption = "F3"
    Change_Command.Caption = "F7 "
           DBGrid1.Columns("cai_number").Caption = "A"
           DBGrid1.Columns("name").Caption = "N"
           DBGrid1.Columns("wna").Caption = "C"
           DBGrid1.Columns("dou").Caption = "D"
           DBGrid1.Columns("jiao").Caption = "J"
           DBGrid1.Columns("mathine").Caption = "M"

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
   Call Del_Win(Hwnd)
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - picButtons.Height - 30
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
With datPrimaryRS
  If .Recordset.RecordCount > 0 Then
        If .Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
            Run_Grid.SetFocus
            Exit Sub
      End If
      temp% = Run_Grid.Columns("Sort_Idx").Value
        i% = temp%

        strsqlrestore = "update fen_pei_table  set  Sort_Idx=Sort_idx+1   where Sort_Idx>=" & temp%
        .Database.Execute strsqlrestore, dbFailOnError
        
        .Recordset.AddNew
        .Recordset.Fields("Sort_idx").Value = temp%
        .Recordset.Fields("mathine").Value = 1
        .Recordset.Update
        .Refresh
        
        .Recordset.FindFirst "Sort_idx=" & temp%
        
        Run_Grid.Refresh
        Run_Grid.AllowAddNew = False
  End If
        Run_Grid.SetFocus
End With
  Exit Sub
ErrorHandler:
    Speak_Error ("须选择配方")
    Run_Grid.SetFocus

End Sub

Private Sub Print_Command_Click()
'   On Error GoTo error_doing
   Dim Y
   Dim i%, k%, Font_Control%
   Dim book   As Variant
   Dim temp(0 To 20) As Single
    Dim Print_Head(1 To 3)

Print_Head(1) = Array("段", "投时", "投时", "投时", "投时", "投时", "投上", "转", "压", "混时", "温", "能", "控关")
Print_Head(2) = Array("  ", "   ", "  ", "油 ", "油  ", "小  ", "    ", "  ", "力", "    ", "  ", "  ", " ")
Print_Head(3) = Array("号", "胶间", "碳间", "一间", "二间", "料间", "料升", "速", "级", "炼间", "度", "量", "制系")


    
   book = Run_Grid.Bookmark
   Screen.MousePointer = vbHourglass
   
    If Printer.ScaleWidth < Run_Grid.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        Run_Grid.SetFocus
        Exit Sub
    End If
  With datPrimaryRS
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
                temp(k) = temp(k - 1) + Run_Grid.Columns(i).Width * (Printer.ScaleWidth) / Run_Grid.Width
                k = k + 1
        End If
   Next i
   End With
    Printer.Print
    Printer.Print
'    Printer.FontName = "Courier"
 '   Printer.FontSize = 20
     'Printer.FontName = "Courier"
    'Printer.FontSize = 18


        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY), 1
        Printer.Print Run_Grid.Caption + space(4) + Date$
    
    'Printer.FontName = "Courier"
    'Printer.FontSize = 12
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
        k% = 0
        For i% = 0 To Run_Grid.Columns.count - 1
            If Run_Grid.Columns(i%).Visible Then
                   Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(Run_Grid.Columns(i%).Caption); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
            End If
        Next i%
        
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
'打印内容：
With datPrimaryRS
        .Refresh
        Do While Not .Recordset.EOF()
            k% = 0
           For i% = 0 To Run_Grid.Columns.count - 1
               If Run_Grid.Columns(i%).Visible = True Then
                     Printer.PSet (temp(k%), Printer.CurrentY)
                     Printer.Print Trim(Run_Grid.Columns(i%).Value); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                     k% = k% + 1
                End If
           Next i%
           Printer.Print String(1, " ")
           .Recordset.MoveNext
         Loop
            Y = Printer.CurrentY + 10
            Printer.Line (0, Y)-(Printer.ScaleWidth, Y)     '  print --------
            Printer.Print
            Printer.EndDoc
        End With
    Screen.MousePointer = vbDefault
'    Run_Grid.Bookmark = book
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
      Run_Grid.Bookmark = book
    

End Sub

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim Single_Weight As Single
Dim kind_code   As Integer

           If ColIndex = 1 Or ColIndex = 2 Then
                If ColIndex = 2 And Check_Float(Run_Grid.Columns(2).Value) = False Then
                        Speak_Error ("必须输入正确斗号")
                        Cancel = True
                        Exit Sub
                 End If
                 Run_Grid.Columns(1).Value = Val(Run_Grid.Columns(1).Value)
                If Run_Grid.Columns(1).Value <> 1 And Run_Grid.Columns(1).Value <> 2 Then
                        Speak_Error ("必须输入正确机组号")
                        Cancel = True
                        Exit Sub
                 End If
                If ColIndex = 2 And Run_Grid.Columns(2).Value <> OldValue Then
                    Cai_Data.Recordset.FindFirst "mathine=" & Run_Grid.Columns(1).Value & " and dou=" & Run_Grid.Columns(2).Value
                    If Cai_Data.Recordset.NoMatch Then
                                Speak_Error ("无此斗号")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                   End If
                    Call DBGrid1_DblClick
                    If Get_Kind(Run_Grid.Columns(1).Value, Run_Grid.Columns(2).Value, _
                          Single_Weight, kind_code) Then
                          Run_Grid.Columns(3).Value = kind_code
                    Else
                                Speak_Error ("斗号无品种号")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                    End If
                    'Run_Grid.Columns(6).Value = CInt(Get_Send_Data(Run_Grid.Columns(1).Value, Run_Grid.Columns(2).Value, Single_Weight))
                    'Run_Grid.SetFocus
                End If
           End If

End Sub

Private Sub Run_Grid_Error(ByVal DataError As Integer, Response As Integer)
        
         Response = 0
End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, Flag As Integer
    Select Case KeyCode
        Case 40
            Flag = 0
            For i = 1 To Run_Grid.Columns.count - 1
              If Run_Grid.Columns(i).Value = "" And Run_Grid.Columns(i).Visible Then
                    Flag = 1
                    Exit For
              End If
            Next i
            If Flag = 0 Then
                Run_Grid.AllowAddNew = True
            Else
                Run_Grid.AllowAddNew = False
            End If
        Case vbKeyF3
            Call Insert_Command_Click
            Run_Grid.SetFocus
            KeyCode = 0
        Case vbKeyF4
            Call Print_Command_Click
            Run_Grid.SetFocus
            KeyCode = 0
        Case vbKeyF6
            Call CmdFind_Click
            Run_Grid.SetFocus
            KeyCode = 0
        Case vbKeyF7
            If cmdRun.Enabled Then
                Call cmdRun_Click
            Else
                Call Change_Command_Click
            End If
            Run_Grid.SetFocus
            KeyCode = 0
        Case vbKeyF10
            Call Del_Command_Click
            KeyCode = 0
            Run_Grid.SetFocus
    End Select
    
    
End Sub

Private Sub run_Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
              If Run_Grid.VisibleRows < 1 Then
                    Exit Sub
              End If
              
            If Run_Grid.col = 3 Or Run_Grid.col = 5 Then
                If LastCol < 3 Then
                        Run_Grid.col = 6
                Else
                        Run_Grid.col = 2
                End If
            End If
            If Run_Grid.Columns(0).Value = "" Then
                    old_index = datPrimaryRS.Recordset.RecordCount + 1
                    Run_Grid.Columns("sort_idx").Value = old_index
                    Run_Grid.Columns("mathine").Value = 1
            End If
            
End Sub

Private Sub Timer1_Timer()
        If Fen_Run_flag = 0 Then
            frmFen_Pei_table!cmdRun.Enabled = True
            frmFen_Pei_table!Change_Command.Enabled = False
            Timer1.Enabled = False
        End If
        
End Sub
'取品种号
Public Function Get_Kind(Mathine As Integer, Dou As Integer, Single_Weight As Single, kind_code As Integer) As Boolean
    Dim dbsNorthwind As Database
    Dim rstTemp As Recordset

    
        Get_Kind = False
        Set dbsNorthwind = OpenDatabase("c:\ljxt\comm.mdb")
        Set rstTemp = dbsNorthwind.OpenRecordset( _
        "SELECT kind_code,single_weight FROM fen_send_table where mathine=" & Mathine & " and dou=" & Dou, dbOpenDynaset, dbReadOnly)
        If Not rstTemp.EOF Then
            Get_Kind = True
            Single_Weight = rstTemp!Single_Weight
            kind_code = rstTemp!kind_code
        End If
        rstTemp.Close
        dbsNorthwind.Close
End Function
'取罐数
Public Function Get_Send_Data(Mathine As Integer, Dou As Integer, Single_Weight As Single) As Integer
    Dim dbsNorthwind As Database
    Dim rstTemp As Recordset

    
            Get_Send_Data = 0
        Set dbsNorthwind = OpenDatabase("c:\ljxt\ljxt.mdb")
        Set rstTemp = dbsNorthwind.OpenRecordset( _
        "SELECT sum(weight) as  mweight FROM tan_hei_table where mathine=" & Mathine & " and dou='" & Dou & "'", dbOpenDynaset, dbReadOnly)
        If Not rstTemp.EOF Then
            If IsNull(rstTemp!mWeight) Then
                Get_Send_Data = 0
            Else
                Get_Send_Data = rstTemp!mWeight / Single_Weight
            End If
        End If
        rstTemp.Close
        dbsNorthwind.Close
End Function

'取罐数
Public Sub Order_Send()
    Dim dbsNorthwind As Database
    Dim rstTemp As Recordset
    Dim rdst As Recordset
    Dim Ord_idx As Integer
    On Error GoTo Exit_Do1
        'datPrimaryRS.Database.Execute "drop table temp_use"
        
        datPrimaryRS.Database.Execute _
         "SELECT ban,cai_number, sum(send_data) as data " _
        & " into  temp_use  from fen_pei_table  group by ban,cai_number ", dbFailOnError
            'order by sum(send_data) desc
        
        
        Set dbsNorthwind = OpenDatabase("c:\ljxt\ljxt.mdb")
        Set rstTemp = dbsNorthwind.OpenRecordset( _
         "SELECT * " _
        & " from temp_use order by ban,data desc ", dbOpenDynaset, dbReadOnly)
        If IsNull(rstTemp!Cai_number) Then
            Call Speak_Error("材料号不能为空")
            GoTo Exit_Do
        End If
        If IsNull(rstTemp!ban) Then
            Call Speak_Error("班号不能为空")
            GoTo Exit_Do
        End If
        
        Ord_idx = 1
        While Not rstTemp.EOF
            Set rdt = dbsNorthwind.OpenRecordset( _
                "SELECT * " _
                & " from fen_pei_table where cai_number ='" & rstTemp!Cai_number & "' and ban=" & rstTemp!ban & " order by send_data desc", dbOpenDynaset)
                While Not rdt.EOF
                    rdt.Edit
                    rdt!Sort_idx = Ord_idx
                    rdt.Update
                    Ord_idx = Ord_idx + 1
                    rdt.MoveNext
                Wend
                rdt.Close
                rstTemp.MoveNext
        Wend
Exit_Do:
        rstTemp.Close
        dbsNorthwind.Close
Exit_Do1:
        datPrimaryRS.Database.Execute "drop table temp_use", dbFailOnError
        
End Sub

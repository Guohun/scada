VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Fen_send_table 
   Caption         =   "风送设备管理"
   ClientHeight    =   6192
   ClientLeft      =   156
   ClientTop       =   936
   ClientWidth     =   9504
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6192
   ScaleWidth      =   9504
   Begin Threed.SSFrame SSFrame1 
      Height          =   1092
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9372
      _Version        =   65536
      _ExtentX        =   16531
      _ExtentY        =   1926
      _StockProps     =   14
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
      ShadowStyle     =   1
      Begin Threed.SSCommand Exit_Command 
         Cancel          =   -1  'True
         Height          =   732
         Left            =   8040
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   1104
         _Version        =   65536
         _ExtentX        =   1940
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
         Font3D          =   2
         Picture         =   "comm_fen.frx":0000
      End
      Begin Threed.SSCommand CmdFind 
         Height          =   732
         Left            =   5280
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   1104
         _Version        =   65536
         _ExtentX        =   1940
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
         Font3D          =   2
         Picture         =   "comm_fen.frx":0452
      End
      Begin Threed.SSCommand Delete_Command 
         Height          =   732
         Left            =   6720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   1104
         _Version        =   65536
         _ExtentX        =   1940
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
         MouseIcon       =   "comm_fen.frx":08A4
         Picture         =   "comm_fen.frx":0CF6
      End
      Begin Threed.SSCommand Print_Command 
         Height          =   732
         Left            =   3600
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   1104
         _Version        =   65536
         _ExtentX        =   1940
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
         Font3D          =   2
         MouseIcon       =   "comm_fen.frx":1148
         Picture         =   "comm_fen.frx":159A
      End
      Begin Threed.SSCommand Insert_Command 
         Height          =   732
         Left            =   1800
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
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
         Picture         =   "comm_fen.frx":19EC
      End
      Begin VB.Label record_prompt 
         Caption         =   "Label1"
         Height          =   372
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1812
      End
   End
   Begin MSDBGrid.DBGrid Run_Grid 
      Bindings        =   "comm_fen.frx":1E3E
      Height          =   4812
      Left            =   120
      OleObjectBlob   =   "comm_fen.frx":1E52
      TabIndex        =   0
      Top             =   1200
      Width           =   9252
   End
   Begin VB.Data Comm_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Fen_send_table"
      Top             =   5640
      Visible         =   0   'False
      Width           =   3132
   End
End
Attribute VB_Name = "Fen_send_table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsDBName As Database
Dim gsRecordSend As Recordset
Dim gsRecordComm As Recordset
'Type send_pei_buffer


'第一次显示操作
Private Sub FirstDo()
Dim i    As Integer
      Run_Grid.Font.Size = 10
      Run_Grid.HeadFont.Size = 12

        For i = 0 To Run_Grid.Columns.count - 1
            If Run_Grid.Columns(i).Visible = True Then
               Run_Grid.Columns(i).Width = 650
            End If
        Next i
'         Run_Grid.Columns("max_sendtime").Width = 800
        Run_Grid.RowHeight = 600
        Run_Grid.HeadLines = 2
      
End Sub







'查找命令
Private Sub CmdFind_Click()
On Error GoTo FindErr
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
  With Comm_Data.Recordset
  
   frmFind.lstFields.AddItem Run_Grid.Columns(2).Caption
   frmFind.lstFields.AddItem Run_Grid.Columns(3).Caption
   frmFind.lstFields.AddItem Run_Grid.Columns(4).Caption
   frmFind.lstFields.AddItem Run_Grid.Columns(5).Caption
   frmFind.lstFields.AddItem Run_Grid.Columns(6).Caption
   frmFind.lstFields.AddItem Run_Grid.Columns(7).Caption

FindStart:
  gbFindFailed = False
  gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOp
  frmFind.txtExpression.Text = gsFindExpr
  frmFind.optFindType(gnFindType) = True

  frmFind.Show vbModal

  
  
 If gbFindFailed = True Then Exit Sub 'find cancelled
 If Run_Grid.Columns(gsFindFiel).DataField <> "name" Then
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
End Sub

Private Sub comm_Data_Reposition()
If China_English = CHINA Then
  If Not Comm_Data.Recordset.EOF And Not Comm_Data.Recordset.BOF Then
    record_prompt.Caption = "当前记录:" & (Comm_Data.Recordset.AbsolutePosition + 1) & " 之 " & Comm_Data.Recordset.RecordCount
  End If
Else

End If

End Sub



Private Sub Delete_command_Click()
  Dim temp_first%, temp%
  If Comm_Data.Recordset.RecordCount < 0 Then
        Exit Sub
  End If
  If Comm_Data.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择记录")
        Exit Sub
   End If

If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
'              temp_first% = DBGrid1(0).FirstRow
'             temp% = DBGrid1(0).row
           Temp_record = Comm_Data.Recordset.AbsolutePosition
           Comm_Data.Recordset.Delete
            Comm_Data.Refresh

           If Temp_record > 1 Then
                Comm_Data.Recordset.AbsolutePosition = Temp_record - 1
           End If
'                If DBGrid1(0).VisibleRows < 1 Then Exit Sub
 '                DBGrid1(0).FirstRow = temp_first%
  '               If temp% < DBGrid1(0).VisibleRows Then
   '                     DBGrid1(0).row = temp%
    '            Else
     '                   DBGrid1(0).row = temp% - 1
      '          End If
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DPei_fan_table.Recordset.Delete
End If
Run_Grid.SetFocus
Run_Grid.Refresh

End Sub

Private Sub Exit_Command_Click()
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
            Call FirstDo
            Comm_Data.DatabaseName = Data_Path & "\ljxt.mdb"
            Comm_Data.RecordSource = "select * from fen_send_table order by name " 'order by  mathine,dou"
            Comm_Data.Refresh
End Sub

Private Sub Form_Paint()
  If China_English = CHINA Then
      Caption = "风送设备管理系统"
      Run_Grid.Columns("name").Caption = "逻辑名"
      Run_Grid.Columns("single_weight").Caption = "单罐重量"
      Run_Grid.Columns("dou").Caption = "斗号"
      Run_Grid.Columns("work_press").Caption = "工作压力"
      
      Run_Grid.Columns("set_high_press").Caption = "中压设定"
      Run_Grid.Columns("set_low_press").Caption = "低压设定"
      
      Run_Grid.Columns("max_sendtime").Caption = "最大输送时间"
      Run_Grid.Columns("congestion_press").Caption = "堵塞高压"
      Run_Grid.Columns("clear_time").Caption = "清扫时间"
      Run_Grid.Columns("mathine").Caption = "机组号"
      Run_Grid.Columns(3).Caption = "品种号"
      Run_Grid.Columns(11).Caption = "停用"
      Run_Grid.Columns(12).Caption = "比例系数"
      Run_Grid.Columns(13).Caption = "备用"
    
    Delete_Command.Caption = "F10删除"
    Insert_Command.Caption = "F3插入"
    CmdFind.Caption = "F6查找"
    Exit_Command.Caption = "Esc退出"
    Print_Command.Caption = "F4打印"
 Else
      Caption = "Sennd "
      Run_Grid.Columns("name").Caption = "logic"
      Run_Grid.Columns("single_weight").Caption = "weight"
      Run_Grid.Columns("dou").Caption = "dou"
      Run_Grid.Columns("work_press").Caption = "power"
      
      Run_Grid.Columns("set_high_press").Caption = "Z"
      Run_Grid.Columns("set_low_press").Caption = "D"
      
      Run_Grid.Columns("max_sendtime").Caption = "T"
      Run_Grid.Columns("congestion_press").Caption = "G"
      Run_Grid.Columns("clear_time").Caption = "Q"
      Run_Grid.Columns("mathine").Caption = "M"
      Run_Grid.Columns(3).Caption = "P"
      Run_Grid.Columns(11).Caption = "T"
      Run_Grid.Columns(12).Caption = "B"
      Run_Grid.Columns(13).Caption = "B"
    
    Delete_Command.Caption = "F10 Del"
    Insert_Command.Caption = "F3 Ins"
    CmdFind.Caption = "F6 Search"
    Exit_Command.Caption = "Esc "
    Print_Command.Caption = "F4 print"

End If

End Sub


Private Sub Form_Terminate()
        Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '      Dim strsqlrestore   As String
'        strsqlrestore = "delete  from   fen_send_table    where is null"
 '       Comm_Data.Database.Execute strsqlrestore, dbFailOnError

Call Del_Win(Hwnd)
End Sub

Private Sub Insert_Command_Click()
Dim temp%
 Dim strsqlrestore   As String
 
 'On Error GoTo ErrorHandler:
 If Run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能插入")
            Run_Grid.SetFocus
            Exit Sub
 End If
With Comm_Data
  If .Recordset.RecordCount > 0 Then
        If .Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
            Run_Grid.SetFocus
            Exit Sub
      End If
      temp% = Run_Grid.Columns("name").Value
        i% = temp%

        strsqlrestore = "update fen_send_table  set  name=name+1   where name>=" & temp%
        .Database.Execute strsqlrestore, dbFailOnError
        
        .Recordset.AddNew
        .Recordset.Fields("name").Value = temp%
        .Recordset.Fields("mathine").Value = 1
        .Recordset.Fields("dou").Value = 0
        .Recordset.Update
        .Refresh
        
        .Recordset.FindFirst "name=" & temp%
        
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
   'On Error GoTo error_doing
   Dim Y, i%, k%
   Dim temp(0 To 20) As Single
    
   Screen.MousePointer = vbHourglass
    If Printer.ScaleWidth < Run_Grid.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        Run_Grid.SetFocus
        Exit Sub
    End If
  With Comm_Data
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
    
    If China_English = CHINA Then
        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
        Printer.Print "设备管理报表 " + Date$
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
    For i% = 0 To Run_Grid.Columns.count - 1
         If Run_Grid.Columns(i%).Visible Then
               Printer.PSet (temp(k%), Printer.CurrentY)
                Printer.Print Run_Grid.Columns(i%).Caption;
                k% = k% + 1
        End If
    Next i%
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    With Comm_Data
    .Recordset.MoveFirst
    j% = 0
    Do While Not .Recordset.EOF()
        If j% < 48 Then
            k% = 0
           For i% = 0 To Run_Grid.Columns.count - 1
              If Run_Grid.Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Run_Grid.Columns(i%).Value; 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
              End If
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
End With

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


Private Sub Run_Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
            Select Case Run_Grid.Columns(ColIndex).DataField
                Case "single_weight", "work_press"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) And KeyAscii <> Asc(".") Then
                        Speak_Error ("必须输入" + Run_Grid.Columns(ColIndex).Caption + "为数值 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                Case "set_high_press", "set_low_press", "send_max_time", "dou", "clear_time", "congestion_press"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入" + Run_Grid.Columns(ColIndex).Caption + "为数字 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                    Case "mathine"
                    If (KeyAscii > Asc("2") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入" + Run_Grid.Columns(ColIndex).Caption + "为数字0,1,2 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                    
            End Select

End Sub

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim temp_max, temp_min As Single
        With Run_Grid
                    Select Case .Columns(ColIndex).DataField
                            Case "single_wight", "work_press"
                              If Not Check_Float(.Columns(ColIndex).Value) Then
                                Speak_Error ("输入 " + .Columns(ColIndex).Caption + "非法")
                                Cancel = True
                                .SetFocus
                           End If
                           Case "set_high_press", "set_low_press"
                              If Not Check_Float(.Columns(ColIndex).Value) Then
                                Speak_Error ("输入 " + .Columns(ColIndex).Caption + "非法")
                                Cancel = True
                                .SetFocus
                             End If
'                        temp_max = IIf(.Columns("set_high_press").Text <> "", .Columns("set_high_press").Value, 0)
 '                       temp_min = IIf(.Columns("set_low_press").Text <> "", .Columns("set_low_press").Value, 0)
  '                  If temp_max < temp_min Then
   '                         Call Speak_Error(.Columns("set_high_press").Caption + "<" + .Columns("set_low_press").Caption)
    '                        .SetFocus
     '                       Cancel = True
'                    End If
                           
                    End Select
End With
End Sub


Private Sub Run_Grid_Error(ByVal DataError As Integer, Response As Integer)
     Debug.Print DataError
     
     Select Case DataError
        ' If database file not found.
        Case 16389
            Call Speak_Error(Run_Grid.Columns("mathine").Caption + "与" + Run_Grid.Columns("dou").Caption + "重")
            Response = vbDataErrContinue
            Run_Grid.SetFocus
            ' Display an Open dialog box.
            'CommonDialog1.ShowOpen
        '...
    End Select

End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i, Flag As Integer
    Select Case KeyCode
        Case 40
            Flag = 0
            For i = 0 To 5
              If Run_Grid.Columns(i).Value = "" And Run_Grid.Columns(i).Visible = True Then
                    Flag = 1
                    Exit For
              End If
            Next i
            If Flag = 0 Then
                Run_Grid.AllowAddNew = True
            Else
                Run_Grid.AllowAddNew = False
            End If
        Case vbKeyF10
            Call Delete_command_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF6
            Call CmdFind_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF3
            Call Insert_Command_Click
            KeyCode = 0
            Run_Grid.SetFocus
            
    End Select
    
        
        
End Sub

Private Sub run_Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
              Dim old_index As Integer
              On Error GoTo errorHandle
              
              If Run_Grid.VisibleRows < 1 Then
                    Exit Sub
              End If
            If Run_Grid.Columns("name").Value = "" Then
                    old_index = Comm_Data.Recordset.RecordCount + 1
                    Run_Grid.Columns("name").Value = old_index
            End If
            Exit Sub
errorHandle:
        Speak_Error ("须选择配方")
              
End Sub


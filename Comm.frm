VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form send_table_input 
   Caption         =   "设备管理"
   ClientHeight    =   5700
   ClientLeft      =   156
   ClientTop       =   936
   ClientWidth     =   8976
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   8976
   Begin Threed.SSFrame SSFrame1 
      Height          =   972
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8652
      _Version        =   65536
      _ExtentX        =   15261
      _ExtentY        =   1714
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
         Left            =   7560
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
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
         Font3D          =   2
         Picture         =   "Comm.frx":0000
      End
      Begin Threed.SSCommand CmdFind 
         Height          =   732
         Left            =   4560
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
         Picture         =   "Comm.frx":0452
      End
      Begin Threed.SSCommand Delete_Command 
         Height          =   732
         Left            =   6000
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
         MouseIcon       =   "Comm.frx":08A4
         Picture         =   "Comm.frx":0CF6
      End
      Begin Threed.SSCommand Print_Command 
         Height          =   732
         Left            =   1440
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
         MouseIcon       =   "Comm.frx":1148
         Picture         =   "Comm.frx":159A
      End
      Begin Threed.SSCommand Send_Command 
         Height          =   732
         Left            =   3000
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   1104
         _Version        =   65536
         _ExtentX        =   1940
         _ExtentY        =   1291
         _StockProps     =   78
         Caption         =   "&Send"
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
         MouseIcon       =   "Comm.frx":19EC
         Picture         =   "Comm.frx":1E3E
      End
      Begin VB.Label record_prompt 
         Caption         =   "Label1"
         Height          =   372
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1812
      End
   End
   Begin VB.Data Comm_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\comm.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "send_table"
      Top             =   6120
      Visible         =   0   'False
      Width           =   3132
   End
   Begin MSDBGrid.DBGrid Run_Grid 
      Bindings        =   "Comm.frx":2290
      Height          =   4452
      Left            =   120
      OleObjectBlob   =   "Comm.frx":22A4
      TabIndex        =   0
      Top             =   1200
      Width           =   8652
   End
End
Attribute VB_Name = "send_table_input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsDBName As Database
Dim gsRecordSend As Recordset
Dim gsRecordComm As Recordset


'Type send_pei_buffer
Private Sub FirstDo()
Dim i    As Integer

      Run_Grid.Columns("time").Visible = False
      Run_Grid.Columns("status").Visible = False
      Run_Grid.Columns("true_data").Visible = False
      Run_Grid.Columns("true_AD").Visible = False
      Run_Grid.Columns("run_state").Visible = False
      
      Run_Grid.Font.Size = 12
      Run_Grid.HeadFont.Size = 12

        For i = 0 To 10
               Run_Grid.Columns(i).Width = 900
        Next i
         Run_Grid.Columns(1).Width = 1200
        Run_Grid.RowHeight = 600
        Run_Grid.HeadLines = 2
      
End Sub






Private Sub CmdFind_Click()
On Error GoTo FindErr
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
  With Comm_Data.Recordset
  
   frmFind.lstFields.AddItem Run_Grid.Columns("name").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("port").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("max").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("min").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("AD_max").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("AD_min").Caption

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
    Exit Sub


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
  
  If Comm_Data.Recordset.EOF() Then
            Speak_Error ("须选择记录")
        Exit Sub
   End If

If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
'              temp_first% = DBGrid1(0).FirstRow
'             temp% = DBGrid1(0).row
           Temp_record = Comm_Data.Recordset.AbsolutePosition
           Comm_Data.Recordset.Delete
           If Temp_record > 1 Then
                Comm_Data.Recordset.AbsolutePosition = Temp_record - 1
           End If
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
On Error Resume Next
Dim i%
            Width = MAIN_MDI.Width  ' Set width of form.
            Height = MAIN_MDI.Height - 600 ' Set height of form.
            Left = 0
            Top = 0
            Call Add_Win(Hwnd)
            Comm_Data.DatabaseName = Data_Path & "\comm.mdb"
            
            Comm_Data.RecordSource = "select * from send_table order by  mathine,code"
            Comm_Data.Refresh
            
            Call FirstDo
            Comm_Data.Database.Execute "DROP TABLE [temp_Send];"
End Sub

Private Sub Form_Paint()
  If China_English = CHINA Then
      Run_Grid.Columns("code").Caption = "代号"
      Run_Grid.Columns("name").Caption = "逻辑名"
      Run_Grid.Columns("min").Caption = "最小值"
      Run_Grid.Columns("max").Caption = "最大值"
      
      Run_Grid.Columns("AD_min").Caption = "最小AD值"
      Run_Grid.Columns("AD_max").Caption = "最大AD值"
      
      Run_Grid.Columns("time").Caption = "已运行时间"
      Run_Grid.Columns("port").Caption = "口号"
      
      Run_Grid.Columns("true_data").Caption = "实际数据"
      Run_Grid.Columns("true_AD").Caption = "实际AD"
      Run_Grid.Columns("status").Caption = "当前状态"
      Run_Grid.Columns("mathine").Caption = "机器号"
      Run_Grid.Columns("run_state").Caption = "运行态"
      
    Delete_Command.Caption = "F10删除"
    Send_Command.Caption = "F5发送"
    CmdFind.Caption = "F6查找"
    Exit_Command.Caption = "Esc退出"
    Print_Command.Caption = "F4打印"
 Else
      Run_Grid.Columns("name").Caption = " name"
      Run_Grid.Columns("min").Caption = "min"
      Run_Grid.Columns("max").Caption = "max"
      Run_Grid.Columns("AD_max").Caption = "MaxAd"
      Run_Grid.Columns("AD_min").Caption = "minAd"
      
      Run_Grid.Columns("time").Caption = "havin"

      Run_Grid.Columns("mathine").Caption = "mathine:"
      Run_Grid.Columns("run_state").Caption = "run_state:"
      
    Send_Command.Caption = "&Send"
    Delete_Command.Caption = "&Delete"
    Exit_Command.Caption = "&Exit"
    Print_Command.Caption = "&Print"

End If

End Sub


Private Sub Form_Terminate()
        Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Dim strsqlrestore   As String
        strsqlrestore = "delete  from   send_table    where [name] is null"
        Comm_Data.Database.Execute strsqlrestore, dbFailOnError

Call Del_Win(Hwnd)
End Sub

Private Sub Print_Command_Click()
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
                Case "max", "min"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) And KeyAscii <> Asc(".") Then
                        Speak_Error ("必须输入" + Run_Grid.Columns(ColIndex).Caption + "为数值 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                Case "AD_max", "AD_min", "part", "mathine"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入" + Run_Grid.Columns(ColIndex).Caption + "为数字 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
            End Select

End Sub

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim i%
Dim rst   As Recordset

Dim temp_max, temp_min As Single
        With Run_Grid
    If ColIndex = 0 Then
        If Right(Trim(.Columns(0).Text), 1) < "0" Or Right(Trim(.Columns(0).Text), 1) > "9" Then
                .Columns(0).Text = .Columns(0).Text & CStr(Comm_Data.Recordset.AbsolutePosition + 1)
        End If
    End If
    If ColIndex > 2 Then
            If .Columns(ColIndex).Text = "" Then
            Call Speak_Error(" 必须输入数据")
            Cancel = True
            .SetFocus
            Exit Sub
            End If
    End If
    If .Columns(ColIndex).DataField = "port" Then     'mathine
            Dim temp_mathine%
            temp_mathine = Val(.Columns(ColIndex).Value)
            If (temp_mathine > 12 Or temp_mathine < 0) Then
                        Call Speak_Error(" 必须输入" + .Columns(ColIndex).Caption + "0~12 ")
                        Cancel = True
                        .SetFocus
                     Exit Sub
            End If
    End If
    If .Columns(ColIndex).DataField = "mathine" Or .Columns(ColIndex).DataField = "port" Then
            Dim temp_port%
            If OldValue = .Columns(ColIndex).Value Then
                Exit Sub
            End If
            temp_port = Val(.Columns("port").Value)
            temp_mathine = Val(.Columns("mathine").Value)
            Comm_Data.Database.Execute "select port,mathine  into  [temp_Send]  from send_table " & _
                    " ", dbFailOnError
                            Set rst = Comm_Data.Database.OpenRecordset("select port,mathine  from  [temp_send]  where mathine=" & temp_mathine & " and  port=" & temp_port)
                            If Not rst.EOF() Then
                                                    rst.Close
                                                    
                                                    Comm_Data.Database.Execute "DROP TABLE [temp_Send];"
                                                                 Call Speak_Error(.Columns(ColIndex).Caption + " 机组口号重")
                                                                  .SetFocus
                                                                 Cancel = True
                                                    Exit Sub
                            End If
                .SetFocus
                rst.Close
                Comm_Data.Database.Execute "DROP TABLE [temp_Send];"
                
                            Set rst = Comm_Data.Database.OpenRecordset("select port,mathine  from  comm_table  where mathine=" & temp_mathine & " and  port=" & temp_port)
                            If rst.EOF() Then
                                                    rst.Close
                                                                 Call Speak_Error("通讯口无此机组和口号重")
                                                                  .SetFocus
                                                                 Cancel = True
                                                    Exit Sub
                            End If
                .SetFocus
                rst.Close
                                
    End If
    If .Columns(ColIndex).DataField = "max" Or .Columns(ColIndex).DataField = "min" Then    'max  min
                      If Not Check_Float(.Columns(ColIndex).Value) Then
                                Speak_Error ("输入 " + .Columns(ColIndex).Caption + "非法")
                                Cancel = True
                                .SetFocus
                                Exit Sub
                        End If

            temp_max = IIf(.Columns("max").Text <> "", .Columns("max").Value, 0)
            temp_min = IIf(.Columns("min").Text <> "", .Columns("min").Value, 0)
       If temp_max < temp_min Then
            Call Speak_Error(.Columns("max").Caption + "<" + .Columns("min").Caption)
            .SetFocus
            Cancel = True
       End If
    End If
End With
End Sub



Private Sub Run_Grid_Error(ByVal DataError As Integer, Response As Integer)
    Debug.Print DataError
    If CHINA = China_English Then
        Select Case Left(Run_Grid.ErrorText, 3)
            Case "The"
                     Speak_Error ("不能输入相同机组的口号")
            Case Else
'                     Speak_Error (Run_Grid.ErrorText)
        End Select
    Else
 '                  Speak_Error (Run_Grid.ErrorText)
    End If
    
    Response = 0
    Run_Grid.SetFocus
End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i, Flag As Integer
    Select Case KeyCode
        Case 40
            Flag = 0
            For i = 0 To Run_Grid.Columns.count - 1
                Debug.Print Run_Grid.Columns(i).Value;
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
        Case vbKeyF6
            Call CmdFind_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF10
            Call Delete_command_Click
            KeyCode = 0
            Run_Grid.SetFocus
    End Select
    
        
        
End Sub

Private Sub run_Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
              Dim old_index As Integer
              If Run_Grid.VisibleRows = 1 Then
                    Exit Sub
              End If
            If Run_Grid.Columns("code").Value = "" Then
                    Run_Grid.Columns("code").Value = 0
                    Run_Grid.Columns("mathine").Value = 1
            End If
            Exit Sub

End Sub


VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form comm_set_Form 
   Caption         =   "Comm_Set_Form"
   ClientHeight    =   5628
   ClientLeft      =   372
   ClientTop       =   516
   ClientWidth     =   8964
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5628
   ScaleWidth      =   8964
   Begin VB.ListBox prompt_list 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      ItemData        =   "comm_set_Form.frx":0000
      Left            =   120
      List            =   "comm_set_Form.frx":0002
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5040
      Width           =   8652
   End
   Begin VB.Data comm_Data 
      Caption         =   "通讯"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\comm.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "comm_table"
      Top             =   5880
      Visible         =   0   'False
      Width           =   4932
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   972
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8616
      _Version        =   65536
      _ExtentX        =   15198
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
         Picture         =   "comm_set_Form.frx":0004
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
         Picture         =   "comm_set_Form.frx":0456
      End
      Begin Threed.SSCommand Delete_Command 
         Height          =   732
         Left            =   6120
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
         MouseIcon       =   "comm_set_Form.frx":08A8
         Picture         =   "comm_set_Form.frx":0CFA
      End
      Begin Threed.SSCommand Print_Command 
         Height          =   732
         Left            =   1680
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
         MouseIcon       =   "comm_set_Form.frx":114C
         Picture         =   "comm_set_Form.frx":159E
      End
      Begin Threed.SSCommand Test_Command 
         Height          =   732
         Left            =   3120
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
         MouseIcon       =   "comm_set_Form.frx":19F0
         Picture         =   "comm_set_Form.frx":1E42
      End
      Begin VB.Label record_prompt 
         Caption         =   "Label1"
         Height          =   372
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1332
      End
   End
   Begin MSDBGrid.DBGrid Run_Grid 
      Bindings        =   "comm_set_Form.frx":2294
      Height          =   3492
      Left            =   120
      OleObjectBlob   =   "comm_set_Form.frx":22A8
      TabIndex        =   8
      Top             =   1200
      Width           =   8616
   End
End
Attribute VB_Name = "comm_set_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub CmdFind_Click()
On Error GoTo FindErr
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
  With Comm_Data.Recordset
  For i = 1 To 4
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
 
  sTmp = "[" + Run_Grid.Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + Trim(gsFindExpr)
  
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
    Else
    
  End If
Else

End If
End Sub











Private Sub comm_Data_Validate(Action As Integer, Save As Integer)
   If Save = -1 Then
   
    End If
End Sub



Private Sub Delete_command_Click()
    If Not Comm_Data.Recordset.EOF() Then
        Comm_Data.Recordset.Delete
        Comm_Data.Recordset.MoveNext
    End If
        Run_Grid.SetFocus
End Sub

Private Sub Exit_Command_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Check_KeyCode(KeyCode, Shift)
End Sub

Public Sub Form_Load()

            Width = MAIN_MDI.Width - 100 ' Set width of form.
            Height = MAIN_MDI.Height - 600 ' Set height of form.
            Left = 0
            Top = 0
    
    change_flag = UNCHANGE
   Comm_Data.DatabaseName = Data_Path & "\comm.mdb"
   Comm_Data.Refresh

            Call Add_Win(Hwnd)
            Call First_Do

End Sub
'初始化标题

Private Sub First_Do()
    If China_English = CHINA Then
        prompt_list.Clear
        prompt_list.AddItem "非专业人员，不得修改"
        Run_Grid.Caption = "多用户卡设置"
        Delete_Command.Caption = "F10删除"
        CmdFind.Caption = "F6查找"
        'print_comm.Caption = "F4查找"
 
        Test_Command.Caption = "&T测试"
        Exit_Command.Caption = "&E退出"
        With Run_Grid
        .Columns(0).Caption = "机组号"
        .Columns(1).Caption = "口地址"
        .Columns(2).Caption = "波特率"
        .Columns(3).Caption = "数据位"
        .Columns(4).Caption = "停止位"
        .Columns(5).Caption = "回送"
        .Columns(6).Caption = "奇偶校验"
        .Columns(7).Caption = "控制"
        .Columns(8).Caption = "有效标志"
        End With
    Else
        prompt_list.AddItem "not  teachnology don't change"
        Delete_Command.Caption = "F10  Del "
        CmdFind.Caption = "F6 search "
        Exit_Command.Caption = "Esc Exit"
        prompt_list.Clear
        prompt_list.AddItem "not change"
        Run_Grid.Caption = " setup"

        Test_Command.Caption = "&T TEst"
        
        With Run_Grid
        .Columns(0).Caption = "name"
        .Columns(1).Caption = "port"
        .Columns(2).Caption = "bounds"
        .Columns(3).Caption = "byte"
        .Columns(4).Caption = "stop"
        .Columns(5).Caption = "echo"
        .Columns(6).Caption = "add"
        .Columns(7).Caption = "control"
        .Columns(8).Caption = "having use"
    End With
    End If
End Sub

Private Sub Form_Paint()
    Call First_Do
    If China_English = CHINA Then
        Caption = "通讯设置"
        Print_Command.Caption = "F4打印"
    Else
        Caption = "comm set"
        Print_Command.Caption = "F4Print"
    End If
    
End Sub










Private Sub Form_Terminate()
            Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If change_flag = CHANGE Then
        If China_English = CHINA Then
            If MsgBox("数据已修改，存盘否", vbOKCancel) = vbYes Then
                Call Ok_Command_Click
            End If
        Else
            If MsgBox("Data have change， Save ", vbOKCancel) = vbYes Then
                Call Ok_Command_Click
            End If
        
        End If
    End If
  Call Del_Win(Hwnd)
End Sub

'口初始化
Public Sub Ok_Command_Click()
        Comm_Data.Recordset.Update
End Sub




'全局变量   comm_port(20)
Public Sub sio_close_do()
   Dim j As Integer
   
   For j = 1 To 20
         If sio_close(j) <> SIO_OK Then
'            If china_english = CHINA Then
'                prompt_list.AddItem "关闭端口" & Comm_Port(j) & "错误" & Chr(13)
'            Else
'                prompt_list.AddItem "Close sio" & Comm_Port(j) & "error " & Chr(13)
'            End If
        End If
     
   Next j

End Sub



Private Sub Print_Command_Click()
   'On Error GoTo error_doing
   Dim Y, i%, k%
   Dim temp(0 To 20) As Single
   
    'Printer.FontName = "黑体"
    'Printer.FontSize = 20
    
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

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    With Run_Grid
    If Not Check_Float(.Columns(ColIndex).Text) Then
            Speak_Error (.Columns(ColIndex).Caption + "非法")
            Cancel = True
            .SetFocus
            Exit Sub
    End If
    
    Select Case ColIndex
        Case 0
            If CInt(.Columns(ColIndex).Text) < 0 Or CInt(.Columns(ColIndex).Text) > 2 Then
                Speak_Error ("机组只能为0，1，2")
                Cancel = True
                .SetFocus
            End If
        Case 1
            If CInt(.Columns(ColIndex).Text) < 1 Or CInt(.Columns(ColIndex).Text) > 10 Then
                Speak_Error ("口号只能为1，2..10")
                Cancel = True
                .SetFocus
                
            End If
        Case 2
            If CInt(.Columns(ColIndex).Text) < 7 Or CInt(.Columns(ColIndex).Text) > 13 Then
                Speak_Error ("波特率只能为7....13")
                Cancel = True
                .SetFocus
                
            End If
        Case 3
            If CInt(.Columns(ColIndex).Text) < 7 Or CInt(.Columns(ColIndex).Text) > 8 Then
                Speak_Error (.Columns(ColIndex).Caption + "只能为7....13")
                Cancel = True
                .SetFocus
                        
            End If
            
        Case 4
                If CInt(.Columns(ColIndex).Text) < 1 Or CInt(.Columns(ColIndex).Text) > 2 Then
                    Speak_Error (.Columns(ColIndex).Caption + "只能为1,2")
                    Cancel = True
                    .SetFocus
                    
                 End If
        
        Case 5
            If CInt(.Columns(ColIndex).Text) < 0 Or CInt(.Columns(ColIndex).Text) > 1 Then
                Speak_Error (.Columns(ColIndex).Caption + "只能为0....1")
                Cancel = True
                .SetFocus
                
            End If
        Case 6
            
            If CInt(.Columns(ColIndex).Text) < 0 Or CInt(.Columns(ColIndex).Text) > 2 Then
                Speak_Error (.Columns(ColIndex).Caption + "只能为 0,1,2")
                Cancel = True
                .SetFocus
                
            End If
            
        Case 7
            If CInt(.Columns(ColIndex).Text) < 0 Or CInt(.Columns(ColIndex).Text) > 2 Then
                Speak_Error (.Columns(ColIndex).Caption + "只能为 0,1,2")
                Cancel = True
                .SetFocus
                
            End If
            
            
        Case 8
            If CInt(.Columns(ColIndex).Text) < 0 Or CInt(.Columns(ColIndex).Text) > 2 Then
                Speak_Error (.Columns(ColIndex).Caption + "只能为 0,1,2")
                Cancel = True
                .SetFocus
                
            End If
    End Select
 End With
End Sub


Private Sub Run_Grid_Error(ByVal DataError As Integer, Response As Integer)
   
   
    Debug.Print DataError
    If CHINA = China_English Then
        Select Case Left(Run_Grid.ErrorText, 3)
            Case "The"
                     Speak_Error ("不能输入相同机组的口号")
            Case Else
                     Speak_Error (Run_Grid.ErrorText)
        End Select
    Else
                   Speak_Error (Run_Grid.ErrorText)
    End If
    
    Response = 0
    Run_Grid.SetFocus
End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i, Flag As Integer
    Select Case KeyCode
'        Case 40
 '           Flag = 0
  '          For i = 0 To 5
   '           If Run_Grid.Columns(i).Value = "" And Run_Grid.Columns(i).Visible = True Then
    '                Flag = 1
     '               Exit For
      '        End If
       '     Next i
        '    If Flag = 0 Then
         '       Run_Grid.AllowAddNew = True
          '  Else
           '     Run_Grid.AllowAddNew = False
            'End If
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
    prompt_list.Clear
   If China_English = CHINA Then
        Select Case Run_Grid.col
        Case 0
                prompt_list.AddItem " 机组号： 0，1，2"
        Case 1
                prompt_list.AddItem "输入段口号：（查看插座） 为 1,2,3等"
        Case 2
                prompt_list.AddItem "7--1200  8--1800 9--2400 10---4800 11--7200  12--9600 13--19200 "
        Case 3
                prompt_list.AddItem "7或8 位"
        Case 4
                prompt_list.AddItem "1或2 位"
        Case 5
                prompt_list.AddItem "0--on 1--off "
        Case 6
                prompt_list.AddItem "0--无 1--奇 2--偶 "
        Case 7
            prompt_list.AddItem "0--无 1--xon/xoff 2--RTS  xon/RTS"
        Case 8
            prompt_list.AddItem "0--有效  1--定时  2--无效"
            
    End Select
End If
End Sub

'use  sio_read  sio_ioctul sio_read  sio_write  test port
Private Sub test_command_Click()
Dim j As Integer
Dim test_str  As String * 256
Dim temp_len  As Long
    prompt_list.Clear
   For j = 0 To 20
     If Comm_Port(j) = -1 Then
        Exit For
     End If
     test_str = "012345678"
     temp_len = sio_write(Comm_Port(j), test_str, 12)
     prompt_list.AddItem "写口 " & Comm_Port(j) & ": " & temp_len & "字节:" & test_str
   Next j
     For j = 0 To 20
         If Comm_Port(j) = -1 Then
            Exit For
         End If
        temp_len = sio_read(Comm_Port(j), test_str, 12)
        If (temp_len > 0) Then
             prompt_list.AddItem "读口 " & Comm_Port(j) & "收到" & temp_len & "字节"
        Else
             prompt_list.AddItem "读口 " & Comm_Port(j) & "失败 "
        End If
     Next j

End Sub


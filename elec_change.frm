VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Elec_Set 
   Caption         =   "Elec_Set"
   ClientHeight    =   5580
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9384
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   9384
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "elec_change.frx":0000
      Height          =   4452
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "elec_change.frx":0013
      TabIndex        =   0
      Tag             =   """0"""
      Top             =   960
      Width           =   4332
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "elec_change.frx":1244
      Height          =   4452
      Index           =   1
      Left            =   4680
      OleObjectBlob   =   "elec_change.frx":1257
      TabIndex        =   1
      Tag             =   """1"""
      Top             =   960
      Width           =   4452
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   7680
      TabIndex        =   2
      TabStop         =   0   'False
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
      Picture         =   "elec_change.frx":2488
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   4680
      TabIndex        =   3
      TabStop         =   0   'False
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
      Picture         =   "elec_change.frx":28DA
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   720
      TabIndex        =   4
      TabStop         =   0   'False
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
      Picture         =   "elec_change.frx":2D2C
   End
   Begin Threed.SSCommand Del_Command 
      Height          =   732
      Left            =   6240
      TabIndex        =   5
      TabStop         =   0   'False
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
      MouseIcon       =   "elec_change.frx":317E
      Picture         =   "elec_change.frx":35D0
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   2640
      TabIndex        =   6
      TabStop         =   0   'False
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
      MouseIcon       =   "elec_change.frx":3A22
      Picture         =   "elec_change.frx":3E74
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Index           =   0
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "node_table"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Index           =   1
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "node_table"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1692
   End
End
Attribute VB_Name = "Elec_Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim change_flag As Boolean
Private Light_Data() As Variant ' 2-dimensional array containing records 临时数
Private Check_Data() As Variant ' 2-dimensional array containing records 临时数
'Private Const MAXCOLS = 3 ' Maximum number of fields in record set.

Private Sub CmdFind_Click()
 Dim iCol, index As Integer
 Dim temp_row As Integer
 On Error Resume Next
 Dim i%
 
index = 0
  Dim sBookMark As String
  Dim sTmp As String
FindStart:
   frmFind.lstFields.AddItem DBGrid1(index).Columns("alias").Caption
  frmFind.lstFields.AddItem DBGrid1(index).Columns("china_prompt").Caption
   frmFind.lstFields.AddItem DBGrid1(index).Columns("english_prompt").Caption
   frmFind.lstFields.AddItem DBGrid1(index).Columns("node_name").Caption
   
   gbFindFailed = False
   'gbFromTableView = False
   frmFind.lstFields.Text = gsFindFiel
   frmFind.lstOperators.Text = gsFindOp
   frmFind.txtExpression.Text = gsFindExpr
   frmFind.optFindType(gnFindType) = True

  frmFind.Show vbModal

  
  
 If gbFindFailed = True Then Exit Sub 'find cancelled
 
    index = 0
    If DBGrid1(index).Columns(gsFindFiel).DataField = "node_name" Then
            sTmp = "[" + DBGrid1(index).Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + Trim(gsFindExpr)
    Else
            sTmp = "[" + DBGrid1(index).Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
    End If
    i = frmFind.lstFields.ListIndex
    sBookMark = Data1(index).Recordset.Bookmark
    Select Case gnFindType
    Case 0
      Data1(index).Recordset.FindFirst sTmp
    Case 1
      Data1(index).Recordset.FindNext sTmp
    Case 2
      Data1(index).Recordset.FindPrevious sTmp
    Case 3
      Data1(index).Recordset.FindLast sTmp
  End Select
  mbNotFound = Data1(index).Recordset.NoMatch
  
  Screen.MousePointer = vbDefault
 If mbNotFound Then
            index = 1
            If DBGrid1(index).Columns(gsFindFiel).DataField = "node_name" Then
                sTmp = "[" + DBGrid1(index).Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + Trim(gsFindExpr)
            Else
                    sTmp = "[" + DBGrid1(index).Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
            End If
            
            i = frmFind.lstFields.ListIndex
            sBookMark = Data1(index).Recordset.Bookmark
            Select Case gnFindType
            Case 0
                Data1(index).Recordset.FindFirst sTmp
            Case 1
                Data1(index).Recordset.FindNext sTmp
                Case 2
                 Data1(index).Recordset.FindPrevious sTmp
                Case 3
                      Data1(index).Recordset.FindLast sTmp
            End Select
               mbNotFound = Data1(index).Recordset.NoMatch
                If mbNotFound = True Then
                         MsgBox "没发现!"
                        Screen.MousePointer = vbDefault
                Else
                    'Data1(index).Recordset.Bookmark = sBookMark
                    sBookMark = Data1(index).Recordset.Bookmark  'save the new position
                End If
  Else
           sBookMark = Data1(index).Recordset.Bookmark  'save the new position

'    Data1(index).Recordset.Bookmark = sBookMark
  End If
  DBGrid1(index).SetFocus
  Exit Sub

FindErr:

    Exit Sub
    mbNotFound = True
    
End Sub

Private Sub DBGrid1_BeforeColUpdate(index As Integer, ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim i%
 Dim rst  As Recordset
 
                       If DBGrid1(index).Columns(ColIndex).Value = "" Then
                                                     Call Speak_Error(DBGrid1(index).Columns(ColIndex).Caption + "不能为空")
                                                     Cancel = True
                                                     DBGrid1(index).SetFocus
                                                    Exit Sub
                      End If

        If ColIndex = 3 Then
                        DBGrid1(index).Columns("in_out").Value = index
                        If DBGrid1(index).Columns(ColIndex).Value = OldValue Then Exit Sub
                       If Not Check_Float(DBGrid1(index).Columns(ColIndex).Value) Then
                                                     Call Speak_Error(DBGrid1(index).Columns(ColIndex).Caption + "范围 0~999")
                                                     Cancel = True
                                                     DBGrid1(index).SetFocus
                                                    Exit Sub
                      End If

                    If DBGrid1(index).Columns(ColIndex).Value < 1 Or DBGrid1(index).Columns(ColIndex).Value > 9999 Then
                                 Call Speak_Error(DBGrid1(index).Columns(ColIndex).Caption + "范围 1~999")
                                 Cancel = True
                                DBGrid1(index).SetFocus
                                 Exit Sub
                                 
                    End If
                Data1(0).Database.Execute "select node_name,in_out  into  [temp_Light]  from node_table ;", dbFailOnError
                Set rst = Data1(0).Database.OpenRecordset("select  * from [temp_Light] where node_name=" & DBGrid1(index).Columns(ColIndex).Value)
                 If Not rst.EOF() Then
                            If rst.Fields("in_out").Value = 0 Then
                                 Call Speak_Error(DBGrid1(index).Columns(ColIndex).Caption + " 与某输入信号重")
                            Else
                                Call Speak_Error(DBGrid1(index).Columns(ColIndex).Caption + " 与某输出信号重")
                            End If
                                 rst.Close
                                 Data1(0).Database.Execute "DROP TABLE [temp_Light];"
                                 Cancel = True
                                 DBGrid1(index).SetFocus
                                 
                                  Exit Sub
                  End If
                rst.Close
                Data1(0).Database.Execute "DROP TABLE [temp_Light];"
        End If 'end colindex


End Sub


Private Sub DBGrid1_Error(index As Integer, ByVal DataError As Integer, Response As Integer)
    Select Case DataError
    Case 3024
            Speak_Error ("没打开数据库")
    Case 16389
         Debug.Print Left(DBGrid1(index).ErrorText, 9)
            Select Case Left(DBGrid1(index).ErrorText, 9)
                Case "The field"
                    Speak_Error ("数据不能为空")
            End Select
    End Select
    Response = 0
    DBGrid1(index).SetFocus
End Sub

Private Sub DBGrid1_GotFocus(index As Integer)
    If index = 0 Then
        DBGrid1(0).Tag = 2
        DBGrid1(1).Tag = 1
    Else
        DBGrid1(1).Tag = 2
        DBGrid1(0).Tag = 0
    End If
        
End Sub

Private Sub DBGrid1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   Dim i, Flag As Integer
    Select Case KeyCode
        Case 40
            Flag = 0
            For i = 0 To DBGrid1(index).Columns.count - 1
              If DBGrid1(index).Columns(i).Value = "" And DBGrid1(index).Columns(i).Visible = True Then
                    Flag = 1
                    Exit For
              End If
            Next i
            If Flag = 0 Then
                DBGrid1(index).AllowAddNew = True
            Else
                DBGrid1(index).AllowAddNew = False
            End If
        Case vbKeyF10
            Call Del_Command_Click
            KeyCode = 0
            DBGrid1(index).SetFocus
        Case vbKeyF3
            Call Insert_Command_Click
            DBGrid1(index).SetFocus
        Case vbKeyF6
            Call CmdFind_Click
            DBGrid1(index).SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            DBGrid1(index).SetFocus
            
    End Select
    
        
        
End Sub

Private Sub DBGrid1_RowColChange(index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    Dim old_index%
    On Error GoTo errorHandle
    
    With DBGrid1(index)
              If .VisibleRows < 1 Then
                    Exit Sub
              End If
              
            If .Columns("sort_idx").Value = "" Then
                    old_index = Data1(index).Recordset.RecordCount + 1
                    .Columns("Sort_Idx").Value = old_index
                    .Columns("in_out").Value = index
            End If
           Exit Sub
errorHandle:
        If index = 0 Then
            If DBGrid1(0).row > 0 Then
                DBGrid1(1).row = DBGrid1(0).row
            End If
            If DBGrid1(0).col > 0 Then
                DBGrid1(1).col = DBGrid1(0).col
            End If
    Else
            
            If DBGrid1(1).row > 0 Then
                DBGrid1(0).row = DBGrid1(1).row
            End If
            If DBGrid1(1).col > 0 Then
                DBGrid1(0).col = DBGrid1(1).col
            End If
    End If
Exit Sub
End With
End Sub


Private Sub Del_Command_Click()
    Dim temp_first As Variant
    Dim temp%
    On Error Resume Next
    If DBGrid1(0).Tag = 2 Then
        If Data1(0).Recordset.EOF() Then Exit Sub
        If Data1(0).Recordset.AbsolutePosition < 0 Then
                Speak_Error ("编辑时不能删除")
                DBGrid1(0).SetFocus
                Exit Sub
        End If
                
        temp% = DBGrid1(0).Columns("sort_idx").Value
        Data1(0).Recordset.Delete
        Data1(0).Refresh
        Dim strsqlrestore   As String
        strsqlrestore = "update   node_table set  sort_idx=sort_idx-1   where sort_idx>" & temp% & "  and  in_out=" & 0
        Data1(0).Database.Execute strsqlrestore, dbFailOnError
        Data1(0).Refresh
        Data1(0).Recordset.FindFirst "sort_idx=" & temp%
         If Data1(0).Recordset.NoMatch Then
            If Not Data1(0).Recordset.EOF() Then
                        Data1(0).Recordset.MoveLast
            End If
         End If
        
    End If
    If DBGrid1(1).Tag = 2 Then
        If Data1(1).Recordset.EOF() Then Exit Sub
        If Data1(1).Recordset.AbsolutePosition < 0 Then
                Speak_Error ("编辑时不能删除")
                DBGrid1(1).SetFocus
                Exit Sub
        End If
        
        temp% = DBGrid1(1).Columns("sort_idx").Value
         Data1(1).Recordset.Delete
         Data1(1).Refresh
 
        strsqlrestore = "update   node_table set  sort_idx=sort_idx-1   where sort_idx>=" & temp% & "  and  in_out=" & 1
        Data1(1).Database.Execute strsqlrestore, dbFailOnError
        Data1(1).Refresh
        Data1(1).Recordset.FindFirst "sort_idx=" & temp%
         If Data1(1).Recordset.NoMatch Then
            If Not Data1(1).Recordset.EOF() Then
                        Data1(1).Recordset.MoveLast
            End If
         End If
         
        
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
Dim i, j As Integer
Width = MAIN_MDI.Width  ' Set width of form.
   Height = MAIN_MDI.Height - 600 ' Set height of form.
     Left = 0
    Top = 0
    On Error Resume Next
Call Add_Win(Hwnd)
change_flag = UNCHANGE


Data1(0).RecordSource = " SELECT *   from node_table where  in_out=0 order by sort_idx"
Data1(1).RecordSource = " SELECT *   from node_table   where  in_out=1 order by sort_idx "
Data1(0).Refresh
Data1(1).Refresh


For j = 0 To 1
    DBGrid1(j).Columns("in_out").Visible = False
    DBGrid1(j).Columns("error_name").Visible = False
    DBGrid1(j).Columns("sort_idx").Visible = False
    For i = 0 To DBGrid1(j).Columns.count - 1
            DBGrid1(j).Columns(i).Width = 800
    Next i
    'DBGrid1(j).Columns("sort_idx").Visible = false
Next j
Data1(0).Database.Execute "DROP TABLE [temp_Light];"
End Sub

Private Sub Form_Paint()

Dim j As Integer

If China_English = CHINA Then
    DBGrid1(0).Caption = "灯/输入信号"
    DBGrid1(1).Caption = "阀/输出信号"
    Del_Command.Caption = "F10删除"
    Exit_Command.Caption = "ESC退出"
    CmdFind.Caption = "F6查找"
    Print_Command.Caption = "F4打印"
  '  Save_Command.Caption = "F5保存"
    Insert_Command.Caption = "F3插入"
    For j = 0 To 1
        DBGrid1(j).Columns(0).Caption = "描述名"
        DBGrid1(j).Columns(1).Caption = "中文名"
        DBGrid1(j).Columns(2).Caption = "英文名"
        DBGrid1(j).Columns(3).Caption = "接点名"
    Next j
Else
    DBGrid1(0).Caption = "Light"
    DBGrid1(1).Caption = "Check"
    Del_Command.Caption = "F10Del"
    Exit_Command.Caption = "&Exit"
    Insert_Command.Caption = "&Save"
    CmdFind.Caption = "F6 search"
    Print_Command.Caption = "F4 print"
    
    For j = 0 To 1
        DBGrid1(j).Columns(0).Caption = "Name"
        DBGrid1(j).Columns(1).Caption = "china name"
        DBGrid1(j).Columns(2).Caption = "eng name"
        DBGrid1(j).Columns(3).Caption = "node name"
    Next j
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call Del_Win(Me.Hwnd)


End Sub

Private Sub Insert_Command_Click()
 Dim iCol, index As Integer
 Dim temp_row As Integer
 Dim i%, temp%
 
 index = -1
If DBGrid1(0).Tag = 2 Then
    index = 0
End If
If DBGrid1(1).Tag = 2 Then
    index = 1
End If
If index = -1 Then Exit Sub
    With DBGrid1(index)
   If .EditActive = True Then
            Speak_Error ("编辑时不能插入")
            .SetFocus
            Exit Sub
    End If
  If Data1(index).Recordset.RecordCount > 0 Then
        If Data1(index).Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
            .SetFocus
        Exit Sub
      End If
      temp% = .Columns("SORT_IDX").Value
        
        Dim strsqlrestore   As String
         strsqlrestore = "update   node_table set  sort_idx=sort_idx+1   where sort_idx>=" & temp% & "  and  in_out=" & index
        Data1(index).Database.Execute strsqlrestore, dbFailOnError
        Data1(index).Refresh
         strsqlrestore = "insert  into  node_table(sort_idx,alias,node_name,in_out)  " _
                                          & "  values (" & temp% & ",0,0," & index & " )"
                                          
        Data1(index).Database.Execute strsqlrestore, dbFailOnError
        
        
        Data1(index).Refresh
        
        Data1(index).Recordset.FindFirst "sort_idx=" & temp%
        
        
        .AllowAddNew = False
  End If
        .SetFocus
        
  Exit Sub
ErrorHandler:
    Speak_Error ("须选择配方")
    .SetFocus
End With

End Sub



Private Sub Print_Command_Click()
   'On Error GoTo error_doing
   Dim Y  As Long
   Dim i%, k%
   Dim temp(0 To 20)     As Long
   Dim book As Variant
   
   book = DBGrid1(0).Bookmark
   
   Screen.MousePointer = vbHourglass
    If Printer.ScaleWidth < DBGrid1(0).Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        DBGrid1(0).SetFocus
        Exit Sub
    End If
  With Data1(0)
        Printer.Height = (.Recordset.RecordCount + 2) * DBGrid1(0).RowHeight
        Printer.Width = Screen.Width
   End With
    
    
   
   Dim index%
   
   temp(0) = 6
   k = 1
   
   For index = 0 To 1
   With DBGrid1(index)
   For i = 0 To .Columns.count - 1
         If .Columns(i).Visible Then
                temp(k) = 0
                temp(k) = temp(k - 1) + .Columns(i).Width * (Printer.ScaleWidth) / .Width
                k = k + 1
        End If
   Next i

   
    Printer.Print
    Printer.Print

        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
        Printer.Print .Caption + space(4) + Date$
   
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
        k% = 0
        For i% = 0 To .Columns.count - 1
            If .Columns(i%).Visible = True Then
                   Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(.Columns(i%).Caption); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
            End If
        Next i%
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
   
   
    Dim j%
    Data1(index).Recordset.MoveFirst
    j% = 0
    Do While Not Data1(index).Recordset.EOF()
        If j% < 48 Then
            k% = 0
           For i% = 0 To .Columns.count - 1
              If .Columns(i%).Visible = True Then
                   Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(.Columns(i%).Value); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
              End If
          Next i%
         Printer.Print String(1, " ")
         Data1(index).Recordset.MoveNext
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
   
    'Do While j% < 48
     '       Printer.Print
      '      j% = j% + 1
    'Loop
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
   Next index
   
'    Printer.Print space(10) + "总数:" & CStr(DCai_liao_table.Recordset.RecordCount)
 '   Printer.Print "第" + CStr(Printer.Page) + "页"
    Printer.EndDoc
    
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
    DBGrid1(0).SetFocus
End Sub

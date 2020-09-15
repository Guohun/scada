VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Cai_liao_table_Input 
   Caption         =   "Cai_Liao_Table_Input"
   ClientHeight    =   5964
   ClientLeft      =   -864
   ClientTop       =   936
   ClientWidth     =   9312
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5964
   ScaleWidth      =   9312
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   730
      Left            =   7440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1330
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
      Font3D          =   2
      Picture         =   "Cai_liao_table_Input.frx":0000
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   4440
      TabIndex        =   5
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
      Font3D          =   2
      Picture         =   "Cai_liao_table_Input.frx":0452
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   1320
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
      Font3D          =   2
      Picture         =   "Cai_liao_table_Input.frx":08A4
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   732
      Left            =   6000
      TabIndex        =   3
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
      Font3D          =   2
      MouseIcon       =   "Cai_liao_table_Input.frx":0CF6
      Picture         =   "Cai_liao_table_Input.frx":1148
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   2880
      TabIndex        =   2
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
      Font3D          =   2
      MouseIcon       =   "Cai_liao_table_Input.frx":159A
      Picture         =   "Cai_liao_table_Input.frx":19EC
   End
   Begin VB.Data DCai_liao_table 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\ljxt.Mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cai_liao_table"
      Top             =   6000
      Visible         =   0   'False
      Width           =   6732
   End
   Begin MSDBGrid.DBGrid run_Grid 
      Bindings        =   "Cai_liao_table_Input.frx":1E3E
      Height          =   3852
      Left            =   840
      OleObjectBlob   =   "Cai_liao_table_Input.frx":1E58
      TabIndex        =   0
      Top             =   960
      Width           =   8052
   End
   Begin VB.Label list_name 
      Caption         =   "list_name"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
End
Attribute VB_Name = "Cai_liao_table_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msBookMark As String
Dim gsDBName As Database
Dim gsRecordSet As Recordset
'如果已有斗号则登记在dou_having
'在run_grid.BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
' 查询已有斗号
Dim Cai_Having(20) As String * 6
Dim Dou_Having(20) As Integer

Private Sub cmdDelete_Click()
    Dim Temp_record As Integer
    
    
    If DCai_liao_table.Recordset.RecordCount <= 0 Then
                    Exit Sub
    End If
    If DCai_liao_table.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择配方")
        Exit Sub
   End If

If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
            Temp_record = DCai_liao_table.Recordset.AbsolutePosition
              DCai_liao_table.Recordset.Delete
              i% = 0
             DCai_liao_table.Refresh
              'DCai_liao_table.Recordset.MoveFirst
              Do While Not DCai_liao_table.Recordset.EOF()
                    Cai_Having(i%) = vFieldVal(DCai_liao_table.Recordset.Fields("cai_number").Value)
                  If Prep_Cai_Liao.wna_code <> 4 Then
                     Dou_Having(i%) = vFieldVal(DCai_liao_table.Recordset.Fields("dou").Value)
                   End If
                
                  DCai_liao_table.Recordset.Edit
                  DCai_liao_table.Recordset.Fields("sort_idx").Value = i% + 1
                  DCai_liao_table.Recordset.Update
                  i% = i% + 1
                  DCai_liao_table.Recordset.MoveNext
            Loop
            Do While i% < 20
                  Cai_Having(i%) = ""
                  Dou_Having(i%) = 0
                  i% = i% + 1
            Loop
            DCai_liao_table.Refresh
           If Temp_record > 1 Then
                DCai_liao_table.Recordset.AbsolutePosition = Temp_record - 1
           End If
           
            Run_Grid.AllowAddNew = True
             
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DPei_fan_table.Recordset.Delete
End If

End Sub











Private Sub CmdFind_Click()
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
FindStart:


  
   
   frmFind.lstFields.AddItem Run_Grid.Columns("cai_number").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("name").Caption
   'frmFind.lstFields.AddItem run_Grid.Columns("").Caption

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
  'Select Case gsFindFiel
   ' Case run_Grid.Columns("_number").Caption
      
    'Case run_Grid.Columns("name").Caption
     ' sTmp = "china_prompt" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
    'Case run_Grid.Columns("m_day").Caption
     ' sTmp = "eng_prompt" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
  'End Select
  i = frmFind.lstFields.ListIndex
  sBookMark = DCai_liao_table.Recordset.Bookmark
  'search for the record
  Select Case gnFindType
    Case 0
      DCai_liao_table.Recordset.FindFirst sTmp
    Case 1
      DCai_liao_table.Recordset.FindNext sTmp
    Case 2
      DCai_liao_table.Recordset.FindPrevious sTmp
    Case 3
      DCai_liao_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = DCai_liao_table.Recordset.NoMatch
  If mbNotFound = True Then MsgBox "not found!"
AfterWhile:
  Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    DCai_liao_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    DCai_liao_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = DCai_liao_table.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    DCai_liao_table.Recordset.Bookmark = sBookMark
  End If

  mbDataChanged = False
  Run_Grid.SetFocus
  
  Exit Sub

FindErr:
    error_name.SetFocus
    Exit Sub
    mbNotFound = True
End Sub

Private Sub DCai_liao_table_Validate(Action As Integer, Save As Integer)
If Save = -1 Then
    For i% = 0 To 5
            If Run_Grid.Columns(i%).Text = "" And Run_Grid.Columns(i%).Visible Then
                Call Speak_Error("应输入" & Run_Grid.Columns(i%).Caption)
                Run_Grid.SetFocus
                Save = 0
                Exit Sub
         End If
   Next i%
End If
End Sub


Private Sub Exit_Command_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    Dim i%
            Width = MAIN_MDI.Width  ' Set width of form.
            Height = MAIN_MDI.Height - 600 ' Set height of form.
            Left = 0
            Top = 0
    For i% = 0 To 20
        Cai_Having(i%) = ""
        Dou_Having(i%) = 0
    Next i%
Call Add_Win(Hwnd)
If China_English = CHINA Then
    Caption = "碳黑输入"
Else
    Caption = "Tan hei "
End If
 DCai_liao_table.DatabaseName = Data_Path & "\ljxt.mdb"

If Prep_Cai_Liao.ShowInTaskbar Then
  DCai_liao_table.RecordSource = "select   *  from cai_liao_table where  mathine= " & Prep_Cai_Liao.mathine_code & _
              " and wna= " & Prep_Cai_Liao.wna_code & "  order by sort_idx "
  
    
Else
  DCai_liao_table.RecordSource = "select   *   from cai_liao_table order by  wna" & "  order by sort_idx "
 End If
 DCai_liao_table.Refresh
  FirstDo
    i% = 0
   If DCai_liao_table.Recordset.RecordCount > 0 Then
     DCai_liao_table.Recordset.MoveFirst
     Do While Not DCai_liao_table.Recordset.EOF
         If vFieldVal(DCai_liao_table.Recordset.Fields(0).Value) <> "" Then
            Cai_Having(i%) = DCai_liao_table.Recordset.Fields(0).Value
            If DCai_liao_table.Recordset.Fields("dou").Value > 0 Then
               Dou_Having(i%) = DCai_liao_table.Recordset.Fields("dou").Value
             End If
            i% = i% + 1
         End If
          DCai_liao_table.Recordset.MoveNext
     Loop
     DCai_liao_table.Recordset.MoveFirst
   End If
 Run_Grid.Columns(3).Visible = False
 If Prep_Cai_Liao.wna_code = 4 Then
          Run_Grid.Columns("jiao").Visible = True
          Run_Grid.Columns("dou").Visible = False
Else
          Run_Grid.Columns("dou").Visible = True
          Run_Grid.Columns("jiao").Visible = False
 End If
 
 
End Sub

Public Sub FirstDo()
        For i = 0 To Run_Grid.Columns.count - 1
               Run_Grid.Columns(i).Width = 1000
        Next i
               Run_Grid.Columns("name").Width = 2000
        Run_Grid.Font.Size = 10
        Run_Grid.HeadFont.Size = 12
        Run_Grid.RowHeight = 600
        Run_Grid.HeadLines = 2
        

End Sub



Private Sub Form_Paint()
If China_English = CHINA Then
  Caption = "原材料表"
  cmdDelete.Caption = "F10删除材料"
  CmdFind.Caption = "F6查询材料"
  Print_Command.Caption = "F4打印"
  Insert_Command.Caption = "F3插入"
  Exit_Command.Caption = "ESC退出"
  
   If Prep_Cai_Liao.wna_code = 4 Then
      list_name.Caption = "1--块胶  2--片胶"
   Else
       list_name.Caption = ""
   End If
    Run_Grid.Caption = Trim(Prep_Cai_Liao.wna_name) & "材料"
    Run_Grid.Columns("cai_number").Caption = "材料编号:"
    Run_Grid.Columns("name").Caption = "材料名称:"
    Run_Grid.Columns("mathine").Caption = "机组号:"
    Run_Grid.Columns("wna").Caption = "秤名:"
    Run_Grid.Columns("dou").Caption = "斗号:"
    Run_Grid.Columns("jiao").Caption = "胶状:"
    Run_Grid.Columns("Sort_Idx").Caption = "目录:"
Else
  Caption = "old_cai_liao"
  cmdDelete.Caption = "&Del"
  CmdFind.Caption = "&Search"
  Exit_Command.Caption = "&Exit"
  Print_Command.Caption = "&Print"
  
  Insert_Command.Caption = "F3 insert"
  
  Run_Grid.Caption = "cai_liao"
   If Prep_Cai_Liao.wna_code = 4 Then
      list_name.Caption = "1--K  2--P"
   Else
       list_name.Caption = ""
   End If
  
  Run_Grid.Columns(0).Caption = "index"
  Run_Grid.Columns(1).Caption = "cai_liao_code:"
  Run_Grid.Columns(2).Caption = "name:"
  Run_Grid.Columns(3).Caption = "mathine:"
  Run_Grid.Columns(4).Caption = "wna:"
  Run_Grid.Columns(5).Caption = "dou:"
  Run_Grid.Columns(6).Caption = "st:"

End If

End Sub


Private Sub Form_Terminate()
    Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Dim strsqlrestore   As String
        strsqlrestore = "delete  from   cai_liao_table    where [cai_number] is null"
        DCai_liao_table.Database.Execute strsqlrestore, dbFailOnError
        Call Del_Win(Hwnd)
End Sub

Private Sub Insert_Command_Click()
 Dim temp%
 'On Error GoTo ErrorHandler:
    If Run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能插入")
            Run_Grid.SetFocus
            Exit Sub
    End If
  If DCai_liao_table.Recordset.RecordCount > 0 Then
        If DCai_liao_table.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
        Exit Sub
      End If
      temp% = Run_Grid.Columns(0).Value
'        run_Grid.AllowAddNew = True
        i% = temp%
        DCai_liao_table.Recordset.FindFirst "sort_idx=" & i%
        Do While Not DCai_liao_table.Recordset.NoMatch
                DCai_liao_table.Recordset.Edit
                DCai_liao_table.Recordset.Fields("sort_idx").Value = i% + 1
                DCai_liao_table.Recordset.Update
                i% = i% + 1
                DCai_liao_table.Recordset.FindNext "sort_idx=" & i%
        Loop
        
        DCai_liao_table.Recordset.AddNew
        DCai_liao_table.Recordset.Fields("sort_idx").Value = temp%
                    DCai_liao_table.Recordset.Fields("wna").Value = Prep_Cai_Liao.wna_code
                    
                    DCai_liao_table.Recordset.Fields("mathine").Value = Prep_Cai_Liao.mathine_code
                     If Prep_Cai_Liao.wna_code = 4 Then
                            DCai_liao_table.Recordset.Fields("jiao").Value = 1
                     End If

        DCai_liao_table.Recordset.Update
        DCai_liao_table.Refresh
        DCai_liao_table.Recordset.FindFirst "sort_idx=" & temp%
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
   'On Error GoTo error_doing
   Dim Y  As Long
   Dim i%, k%
   Dim temp(0 To 20)     As Long
   Dim book   As Variant
   book = Run_Grid.Bookmark
   
   Screen.MousePointer = vbHourglass
    If Printer.ScaleWidth < Run_Grid.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        Run_Grid.SetFocus
        Exit Sub
    End If
  With DCai_liao_table
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
    
'    Printer.NewPage

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
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(Run_Grid.Columns(i%).Caption); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                    
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
                   Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(Run_Grid.Columns(i%).Caption); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
            End If
        Next i%
    
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    DCai_liao_table.Recordset.MoveFirst
    j% = 0
    Do While Not DCai_liao_table.Recordset.EOF()
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
         DCai_liao_table.Recordset.MoveNext
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
        
    'Do While j% < 48
     '       Printer.Print
      '      j% = j% + 1
    'Loop
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
'    Printer.Print space(10) + "总数:" & CStr(DCai_liao_table.Recordset.RecordCount)
 '   Printer.Print "第" + CStr(Printer.Page) + "页"
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    Run_Grid.Bookmark = book
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
      Run_Grid.Bookmark = book
End Sub






Private Sub run_Grid_AfterColEdit(ByVal ColIndex As Integer)
            Select Case ColIndex
                Case 1
                    If Len(Run_Grid.Columns("cai_number").Text) > 6 Then
                        Speak_Error ("材料编号长度必须《6")
                        Run_Grid.Columns("cai_number").Text = Left(Run_Grid.Columns("cai_number").Text, 6)
                    End If
                Case 2
                    If Len(Run_Grid.Columns("name").Text) > 15 Then
                        Speak_Error ("材料名称长度必须《16")
                        Run_Grid.Columns("name").Text = Left(Run_Grid.Columns("name").Text, 15)
                    End If
                    
                Case 3
                    If Run_Grid.Columns("mathine").Text = "" Then
                        Run_Grid.Columns("mathine").Value = 1
                    End If
                    If Run_Grid.Columns("mathine").Value > 2 Then
                        Speak_Error ("必须输入机组0~2")
                        Run_Grid.Columns("mathine").Value = 1
                    End If
                Case 5
                    If (Run_Grid.Columns("dou").Text = "") And Run_Grid.Columns("dou").Visible = True Then
                        Run_Grid.Columns("dou").Value = 1
                        Run_Grid.SetFocus
                    End If
            End Select
            
            
End Sub




Private Sub Run_Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
            
            Dim i%
            Select Case ColIndex
                Case 3, 5
                
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入" + Run_Grid.Columns(ColIndex).Caption + "范围 1~9")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                Case 6
                    If KeyAscii >= Asc("3") Or KeyAscii < Asc("1") Then
                        Call Speak_Error("必须输入胶1~2")
                        Cancel = True
                        Run_Grid.SetFocus
                        
                     End If
                End Select
            
End Sub

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim i%, temp_dou%, temp_local%
    If ColIndex = 1 Then
            If Run_Grid.Columns("cai_number").Value = "" Then
                    Speak_Error ("必须输入配方编号")
                    Cancel = True
                    Exit Sub
             End If
            'temp_local% = DCai_liao_table.Recordset.AbsolutePosition
            ' For i% = 0 To 20
             '   If Trim(Cai_Having(i%)) = "" Then
              '          Exit For
              '  End If
               ' If Trim(Cai_Having(i%)) = run_Grid.Columns(1).Value And i% <> temp_local% Then
               '     Speak_Error (run_Grid.Columns(1).Caption & "重")
             '       Cancel = True
             '       Exit Sub
              '  End If
             ' Next i%
            'If temp_local% < 0 Then
             '   temp_local% = 0
           ' End If
            'Cai_Having(temp_local%) = run_Grid.Columns(1).Value
    End If
    If ColIndex = 5 Then            'dou
            With Run_Grid
                If Check_Float(.Columns("dou").Value) = False Then
                                Speak_Error ("斗号非法")
                                Cancel = True
                                .SetFocus
                                 Exit Sub
                End If
                  If .Columns("dou").Value < 0 Or .Columns("dou").Value > 10 Then
                                Speak_Error ("斗号不能大于10")
                                Cancel = True
                                .SetFocus
                                
                                 Exit Sub
                End If
            End With
            temp_local% = DCai_liao_table.Recordset.AbsolutePosition
             For i% = 0 To 20
                If Dou_Having(i%) = 0 Then
                        Exit For
                End If
                If Dou_Having(i%) = Run_Grid.Columns("dou").Value And i% <> temp_local% Then
                    Speak_Error (Run_Grid.Columns("dou").Caption & "重")
                    Cancel = True
                    Run_Grid.SetFocus
                    Exit Sub
                End If
            Next i%
             If temp_local% < 0 Then
                 temp_local% = 0
             End If
          
             Dou_Having(temp_local%) = Run_Grid.Columns("dou").Value
    End If
    If ColIndex = 6 Then              'jiao
            With Run_Grid
                If Check_Float(.Columns(ColIndex).Value) = False Then
                                Speak_Error (.Columns(ColIndex).Caption & "非法")
                                Cancel = True
                                .SetFocus
                                 Exit Sub
                End If
                
                If .Columns(ColIndex).Value > 2 Or .Columns(ColIndex).Value < 1 Then
                    Speak_Error (.Columns(ColIndex).Caption & "只能为1,2")
                    Cancel = True
                    .SetFocus
                End If
            End With
    End If
End Sub


Private Sub run_Grid_GotFocus()
If Run_Grid.col = 0 Then
    Run_Grid.col = 1
End If
End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, Flag As Integer
    Select Case KeyCode
        Case 40
            Flag = 0
            For i = 0 To Run_Grid.Columns.count - 1
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
         Case vbKeyEscape
                Unload Me
        Case vbKeyF10
            Call cmdDelete_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF3
            Call Insert_Command_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF6
            Call CmdFind_Click
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
        If Run_Grid.col = 0 Then
                Run_Grid.col = LastCol
        End If
              
            If Run_Grid.Columns("wna").Value = "" Then
                    old_index = DCai_liao_table.Recordset.RecordCount + 1
                    Run_Grid.Columns("Sort_Idx").Value = old_index
                    Run_Grid.Columns("wna").Value = Prep_Cai_Liao.wna_code
                    
                    Run_Grid.Columns("mathine").Value = Prep_Cai_Liao.mathine_code
                     If Prep_Cai_Liao.wna_code = 4 Then
                            Run_Grid.Columns("jiao").Value = 1
                     End If
            End If
            Exit Sub
errorHandle:
        Speak_Error ("须选择配方")
End Sub

VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Tan_hei_table_Input 
   Caption         =   "tan_hei"
   ClientHeight    =   5292
   ClientLeft      =   12
   ClientTop       =   900
   ClientWidth     =   9264
   KeyPreview      =   -1  'True
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5292
   ScaleWidth      =   9264
   Begin VB.Data DTan_hei_table 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\ljxt.Mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tan_hei_table"
      Top             =   4560
      Visible         =   0   'False
      Width           =   912
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "tan_hei_table.frx":0000
      Height          =   1500
      Left            =   120
      OleObjectBlob   =   "tan_hei_table.frx":0013
      TabIndex        =   1
      Top             =   3600
      Width           =   9132
   End
   Begin MSDBGrid.DBGrid run_Grid 
      Bindings        =   "tan_hei_table.frx":105D
      Height          =   2532
      Left            =   120
      OleObjectBlob   =   "tan_hei_table.frx":1076
      TabIndex        =   0
      Top             =   960
      Width           =   9132
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   7680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
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
      Picture         =   "tan_hei_table.frx":2FC5
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   4560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
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
      Picture         =   "tan_hei_table.frx":3417
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   1440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
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
      Picture         =   "tan_hei_table.frx":3869
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   732
      Left            =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
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
      MouseIcon       =   "tan_hei_table.frx":3CBB
      Picture         =   "tan_hei_table.frx":410D
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   3000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
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
      MouseIcon       =   "tan_hei_table.frx":455F
      Picture         =   "tan_hei_table.frx":49B1
   End
   Begin VB.Data Cai_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cai_liao_table"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.Label list_name 
      Caption         =   "list_name"
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "Tan_hei_table_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msTBookMark As String
Dim Tan_Ban_weight As Single
Dim Tan_Weight(0 To 40)  As Single
Dim change_flag  As Boolean

Public Sub Add_list()
    Dim number_str    As String * 8, name_str As String * 15
     Dim i%
    Dim Total_Weight As Single
    Total_Weight = 0
    i = 0
    Tan_Weight(i) = 0
    change_flag = UNCHANGE
    
     DTan_hei_table.Refresh
     If DTan_hei_table.Recordset.RecordCount > 0 Then
         Do While Not DTan_hei_table.Recordset.EOF
                        
                        If Not IsNull(DTan_hei_table.Recordset.Fields("weight").Value) Then
                            Tan_Weight(i) = DTan_hei_table.Recordset.Fields("weight").Value
                        Else
                            Tan_Weight(i) = 0
                        End If
                        Total_Weight = Total_Weight + Tan_Weight(i)
                        
                        DTan_hei_table.Recordset.MoveNext
                        i = i + 1
                        Tan_Weight(i) = 0

        Loop
        DTan_hei_table.Refresh
    End If
    If Total_Weight > Tan_Ban_weight Then
                Speak_Error ("超重: '" & Total_Weight - Tan_Ban_weight & "'公斤")
    End If
End Sub















Private Sub cmdDelete_Click()
    Dim Temp_record As Integer
    If DTan_hei_table.Recordset.RecordCount <= 0 Then
                    Exit Sub
    End If
    If DTan_hei_table.Recordset.EOF < 0 Then
            Speak_Error ("须选择配方")
            Run_Grid.SetFocus
        Exit Sub
   End If
If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
          Temp_record = DTan_hei_table.Recordset.Fields("sort_idx").Value
          Run_Grid.AllowAddNew = False
           DTan_hei_table.Recordset.Delete
           i% = 1
           DTan_hei_table.Recordset.MoveFirst
              Do While Not DTan_hei_table.Recordset.EOF()
               If IsNull(DTan_hei_table.Recordset.Fields("weight").Value) Then
                   Tan_Weight(i% - 1) = 0
                Else
                    Tan_Weight(i% - 1) = DTan_hei_table.Recordset.Fields("weight").Value
                End If
                  DTan_hei_table.Recordset.Edit
                  DTan_hei_table.Recordset.Fields("sort_idx").Value = i%
                  DTan_hei_table.Recordset.Update
                  i% = i% + 1
                  DTan_hei_table.Recordset.MoveNext
            Loop
            '.Refresh
            With DTan_hei_table.Recordset
                .FindFirst "sort_idx=" & Temp_record
                If .NoMatch Then
                    If .RecordCount > 0 Then
                     .MoveLast
                    End If
                End If
            End With
            
            Run_Grid.SetFocus
            Run_Grid.AllowAddNew = True
           Run_Grid.AllowAddNew = True
    
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DTan_hei_table.Recordset.Delete
End If
End Sub


Private Sub CmdFind_Click()
 On Error GoTo FindErr
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
  'load the column names into the find form
FindStart:

 
  
   frmFind.lstFields.AddItem Run_Grid.Columns("cai_number").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("name").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("weight").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("dou").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("drop_do").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("fast_do").Caption
      
  '
    'reset the flags
  gbFindFailed = False
  gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOpen
  frmFind.txtExpression.Text = gsFindExpror
  frmFind.optFindType(gsFindType) = True
'  mbNotFound = False
  frmFind.Show vbModal
  'If gbFindFailed = True Then   'find cancelled
  '  GoTo AfterWhile
  'End If
  
  
  If gbFindFailed = True Then Exit Sub
  
   If DTan_hei_table.Recordset.Fields(Run_Grid.Columns(gsFindFiel).DataField).Type = dbText Then
        sTmp = "[" + Run_Grid.Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
    Else
        sTmp = "[" + Run_Grid.Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + Trim(gsFindExpr)
   End If

  
  i = frmFind.lstFields.ListIndex
  sBookMark = DTan_hei_table.Recordset.Bookmark
  'search for the record
  Select Case gsFindType
    Case 0
      DTan_hei_table.Recordset.FindFirst sTmp
    Case 1
      DTan_hei_table.Recordset.FindNext sTmp
    Case 2
      DTan_hei_table.Recordset.FindPrevious sTmp
    Case 3
      DTan_hei_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = DTan_hei_table.Recordset.NoMatch
  If mbNotFound = True Then MsgBox "not found!"
AfterWhile:

  Screen.MousePointer = vbDefault

  If gbFindFaile = True Then   'go back to original row
    DTan_hei_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    DTan_hei_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = DTan_hei_table.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    DTan_hei_table.Recordset.Bookmark = sBookMark
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











Private Sub DBGrid1_DblClick()
                    Run_Grid.Columns("cai_number").Text = DBGrid1.Columns("cai_number").Text
                    Run_Grid.Columns("name").Text = DBGrid1.Columns("name").Text
                    Run_Grid.Columns("dou").Text = DBGrid1.Columns("dou").Text
                    Run_Grid.Columns("wna").Text = DBGrid1.Columns("wna").Text
                    Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                    Run_Grid.SetFocus
                    
End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
                   Dim i%
           If KeyCode = vbKeyTab Then
            'run_Grid.SetFocus
                Exit Sub
        End If
     If KeyCode = 13 Then
                    Run_Grid.Columns("cai_number").Text = DBGrid1.Columns("cai_number").Text
                    Run_Grid.Columns("name").Text = DBGrid1.Columns("name").Text
                    Run_Grid.Columns("dou").Text = DBGrid1.Columns("dou").Text
                    Run_Grid.Columns("wna").Text = DBGrid1.Columns("wna").Text
                    Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                    DBGrid1.SetFocus
                    Run_Grid.col = 5
  End If
 If (KeyCode <> 40 And KeyCode <> 38) Then KeyCode = 0
End Sub



Private Sub DTan_hei_table_Error(DataErr As Integer, Response As Integer)
        Speak_Error ("数据空件错误代号:" & Error$(DataErr))
        Response = 0  'Throw away the error
End Sub



Private Sub Exit_Command_Click()
    Unload Me
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
'On Error Resume Next
     Width = MAIN_MDI.Width  ' Set width of form.
       Height = MAIN_MDI.Height + 400
       Left = 0
      Top = 0
 Call Add_Win(Hwnd)


' run_Grid.SetFocus
 Cai_Data.DatabaseName = Data_Path & "\ljxt.mdb"
 DTan_hei_table.DatabaseName = Data_Path & "\ljxt.mdb"

 DTan_hei_table.RecordSource = "select * from tan_hei_table where pei_number like '" & Pei_fan_table_Input.Pei_Fan_Number & "'   and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & " order by sort_idx "
 DTan_hei_table.Refresh
 
 Cai_Data.RecordSource = "select * from cai_liao_table  where  wna=1 and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & "  and  cai_number <> '  '   " & " order by dou"
 Cai_Data.Refresh
 DBGrid1.Refresh
 Run_Grid.Caption = Pei_fan_table_Input.Pei_Fan_Number & "炭黑配方"
 
 If DTan_hei_table.Recordset.RecordCount = 0 Then
     Run_Grid.Columns("pei_number").Value = Pei_fan_table_Input.Pei_Fan_Number             'index
     Run_Grid.Columns("add_time").Value = 1
     Run_Grid.Columns("Sort_idx").Value = 1
 End If
 Call FirstDo
 'Call Add_list
End Sub

Public Sub FirstDo()
        Dim i%, count As Integer
        Dim gsDBName As Database
        Dim gsThisSet As Recordset

        Set gsDBName = OpenDatabase("c:\ljxt\COMM.mdb")
        Set gsThisSet = gsDBName.OpenRecordset( _
            "select code ,name,max  from  send_table   Where mathine = " & Pei_fan_table_Input.Pei_Fan_Mathine, dbOpenDynaset)
            Do While Not gsThisSet.EOF
                If Not IsEmpty(gsThisSet.Fields(0).Value) Then
                        If gsThisSet.Fields(0).Value = 1 Then
                            Tan_Ban_weight = gsThisSet.Fields(2).Value
                        Exit Do
                        End If
                End If
                gsThisSet.MoveNext
             Loop
        gsThisSet.Close
        gsDBName.Close
        
        For i = 0 To Run_Grid.Columns.count - 1
               Run_Grid.Columns(i).Width = 700
        Next i
        Run_Grid.Columns("password").Visible = False
        Run_Grid.Columns("name").Width = 1600
        Run_Grid.Columns("pei_number").Visible = False
        Run_Grid.Columns("wna").Visible = False
        Run_Grid.Columns("mathine").Visible = False
        'run_Grid.Font.Size = 10
        'run_Grid.HeadFont.Size = 12
        'run_Grid.RowHeight = 600
        'run_Grid.HeadLines = 2
                
        DBGrid1.Columns("jiao").Visible = False
        'DBGrid1.Font.Size = 10
        'DBGrid1.HeadFont.Size = 10
        'DBGrid1.RowHeight = 300
        'DBGrid1.HeadLines = 2
        
End Sub


Private Sub Form_Paint()
        If China_English = CHINA Then
          Caption = "碳黑录入"
          
          cmdDelete.Caption = "F10删除记录"
          Exit_Command.Caption = "ESC退出"
          Print_Command.Caption = "F4打印"
          Insert_Command.Caption = "F3插入"
          CmdFind.Caption = "F6查找"
          Run_Grid.Columns("sort_idx").Caption = "目录"
          Run_Grid.Columns("pei_number").Caption = "配方编号:"
          Run_Grid.Columns("cai_number").Caption = "材料编号:"
          Run_Grid.Columns("name").Caption = "材料名称:"
          Run_Grid.Columns("dou").Caption = "斗号:"
          Run_Grid.Columns("weight").Caption = "所需重量:"
          Run_Grid.Columns("fast_do").Caption = "快慢提前量:"
          Run_Grid.Columns("drop_do").Caption = "点动提前量:"
          Run_Grid.Columns("gon_cha").Caption = "允许公差:"
          Run_Grid.Columns("control_time").Caption = "控制时间:"
          Run_Grid.Columns("add_time").Caption = "加料次数："
          Run_Grid.Columns("stop_time").Caption = "稳定时间："
          Run_Grid.Columns("wna").Caption = "秤"
           DBGrid1.Columns("cai_number").Caption = "材料编号"
           DBGrid1.Columns("name").Caption = "材料名称"
           DBGrid1.Columns("wna").Caption = "秤名"
           DBGrid1.Columns("dou").Caption = "斗名"
           DBGrid1.Columns("jiao").Caption = "胶状"
           DBGrid1.Columns("mathine").Caption = "机组"
           list_name.Caption = "用TAB键盘切换 表格"

        Else
        
          Caption = "tan_hei_input"

          
          cmdDelete.Caption = "F10 Del"
          Exit_Command.Caption = "ESC"
          Print_Command.Caption = "F4Print"
          Insert_Command.Caption = "F3 insert"
          CmdFind.Caption = "F6 search"
          
        Run_Grid.Columns("sort_idx").Caption = "S"
          Run_Grid.Columns(1).Caption = "A:"
          Run_Grid.Columns(2).Caption = "B:"
          Run_Grid.Columns(3).Caption = "C:"
          Run_Grid.Columns(4).Caption = "D:"
          Run_Grid.Columns(5).Caption = "E:"
          Run_Grid.Columns(6).Caption = "F:"
          Run_Grid.Columns(7).Caption = "G:"
          Run_Grid.Columns(8).Caption = "H:"
          Run_Grid.Columns(9).Caption = "control_time"
          Run_Grid.Columns(10).Caption = "add_time"
          Run_Grid.Columns(11).Caption = "stop_time"
          Run_Grid.Columns(12).Caption = "wna"

          DBGrid1.Columns("cai_number").Caption = "code"
          DBGrid1.Columns("name").Caption = "name"
           DBGrid1.Columns("mathine").Caption = "mathine"
           DBGrid1.Columns("jiao").Caption = "jiao"
           DBGrid1.Columns("mathine").Caption = "mathine"
           DBGrid1.Columns("wna").Caption = "che"
           DBGrid1.Columns("dou").Caption = "dou"
           
           list_name.Caption = "use TAB key change grid"
        End If

End Sub









Private Sub Form_Terminate()
Call Del_Win(Hwnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
       On Error Resume Next
        DTan_hei_table.UpdateControls
        strsqlrestore = "delete  from   tan_hei_table    where [cai_number] is null  or weight is null"
        DTan_hei_table.Database.Execute strsqlrestore, dbFailOnError
       Call Del_Win(Hwnd)

End Sub







Private Sub Insert_Command_Click()
Dim temp%
 On Error GoTo ErrorHandler:
  If Run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能插入")
            Run_Grid.SetFocus
            Exit Sub
 End If

  If DTan_hei_table.Recordset.RecordCount > 0 Then
        If DTan_hei_table.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
        Exit Sub
      End If
      temp% = Run_Grid.Columns(0).Value

        i% = temp%
        DTan_hei_table.Recordset.FindFirst "sort_idx=" & i%
        Do While Not DTan_hei_table.Recordset.NoMatch
                DTan_hei_table.Recordset.Edit
                DTan_hei_table.Recordset.Fields("sort_idx").Value = i% + 1
                DTan_hei_table.Recordset.Update
                i% = i% + 1
                DTan_hei_table.Recordset.FindNext "sort_idx=" & i%
        Loop
        
        Dim strsqlrestore   As String
         strsqlrestore = "insert  into tan_hei_table (sort_idx,pei_number,add_time)  " _
                                          & "  values (" & temp% & ",'" & Pei_fan_table_Input.Pei_Fan_Number & "'," & 1 & ")"
                                          
        DTan_hei_table.Database.Execute strsqlrestore, dbFailOnError
        
        DTan_hei_table.Refresh
        DTan_hei_table.Recordset.FindFirst "sort_idx=" & temp%
        'Run_Grid.Refresh
        
        Run_Grid.AllowAddNew = False
  End If
        Run_Grid.SetFocus
        
  Exit Sub
ErrorHandler:
 '   Speak_Error ("Input error")
    Run_Grid.SetFocus

End Sub


Private Sub Print_Command_Click()
'   On Error GoTo error_doing
Dim Print_Head(3)

Print_Head(1) = Array("目", "材 编", "材 名", "斗", "加次", "所重", "快前", "点前", "允公", "控时", "稳时")
Print_Head(2) = Array(" ", "   ", "   ", "  ", "  ", "  ", "慢 ", "动 ", "    ", "  ", "   ")
Print_Head(3) = Array("录", "料 号", "料 称", "号", "料数", "须量", "提量", "提量", "许差", "制间", "定间")

   Dim Y  As Long
   Dim i%, k%, Font_Control
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
  With DTan_hei_table
        Printer.Height = (.Recordset.RecordCount + 8) * Run_Grid.RowHeight
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
 '   Printer.FontSize = 18
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
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
  '  Printer.FontName = "Courier"
   ' Printer.FontSize = 12
    With Run_Grid
     For Font_Control = 1 To 3
        k% = 0
        For i = LBound(Print_Head(1)) To UBound(Print_Head(1))
                Printer.PSet (temp(k%), Printer.CurrentY)
                Printer.Print Print_Head(Font_Control)(i);
                k% = k% + 1
        Next i
        Printer.Print String(1, " ")
    Next Font_Control
    End With
    Erase Print_Head
    
           
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    DTan_hei_table.Recordset.MoveFirst
    j% = 0
    Do While Not DTan_hei_table.Recordset.EOF()
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
         DTan_hei_table.Recordset.MoveNext
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
'    Printer.Print Space(10) + "总数:" & CStr(DTan_hei_table.Recordset.RecordCount)
    'Printer.Print "第" + CStr(Printer.Page) + "页"
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    Run_Grid.Bookmark = book
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
      Run_Grid.Bookmark = book
    
End Sub


Private Sub run_Grid_AfterUpdate()
            change_flag = UNCHANGE
End Sub

Private Sub Run_Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
            change_flag = CHANGE
            Select Case Run_Grid.Columns(ColIndex).DataField
                Case "weight", "fast_do", "drop_do", "gon_cha", "control_time"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) And KeyAscii <> Asc(".") Then
                        Speak_Error ("必须输入重量为数值 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                    
                Case "add_time", "dou", "stop_time"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入为数字 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                    
            End Select
            
End Sub

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
                    Dim i%
                    Dim Total_Weight As Single
            
            Select Case Run_Grid.Columns(ColIndex).DataField
                Case "cai_number"
                        If Run_Grid.Columns("cai_number").Value = "" Then
                                     Speak_Error ("必须输入材料编号")
                                      Cancel = True
                                      Run_Grid.SetFocus
                                      Exit Sub
                        End If
                        If Run_Grid.Columns("cai_number").Value <> OldValue Then
                                Cai_Data.Recordset.FindFirst "cai_number=  '" & Run_Grid.Columns("cai_number").Value & "'"
                                If Cai_Data.Recordset.NoMatch Then
                                    Speak_Error ("无此材料编号")
                                    Run_Grid.SetFocus
                                    Cancel = True
                                    Exit Sub
                   End If
                    Call DBGrid1_DblClick
                    Run_Grid.SetFocus
                End If
                Case "dou"
                        If Check_Float(Run_Grid.Columns("dou").Value) = False Then
                                     Speak_Error ("必须输入正确斗号")
                                      Cancel = True
                                      Run_Grid.SetFocus
                                      Exit Sub
                        End If
                        If Run_Grid.Columns("dou").Value <> OldValue Then
                                Cai_Data.Recordset.FindFirst "dou=  " & Run_Grid.Columns("dou").Value
                                If Cai_Data.Recordset.NoMatch Then
                                    Speak_Error ("无此斗号")
                                    Cancel = True
                                    Run_Grid.SetFocus
                                Exit Sub
                   End If
                            Call DBGrid1_DblClick
                            Run_Grid.SetFocus
                End If
                Case "gon_cha", "stop_time", "control_time", "drop_do", "fast_do"
                         'If Run_Grid.Columns(ColIndex).Value = "" Then
                          '           Speak_Error ("必须输入" + Run_Grid.Columns(ColIndex).Caption)
                           '           Cancel = True
                            '          Run_Grid.SetFocus
                        'End If

                        If Not Check_Float(Run_Grid.Columns(ColIndex).Value) Then
                                Speak_Error ("输入 " + Run_Grid.Columns(ColIndex).Caption + "非法")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                        End If

                        Run_Grid.Columns(ColIndex).Value = CSng(Run_Grid.Columns(ColIndex).Value)
                        Debug.Print Run_Grid.Columns(ColIndex).Value
                        If IsNull(Run_Grid.Columns(ColIndex).Value) Then
                                     Speak_Error ("必须输入" + Run_Grid.Columns(ColIndex).Caption)
                                      Cancel = True
                                      Run_Grid.SetFocus
                        End If

                Case "weight"
                         If Not Check_Float(Run_Grid.Columns(ColIndex).Value) Then
                                Speak_Error ("输入 " + Run_Grid.Columns(ColIndex).Caption + "非法")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                        End If
                         If Not Check_Float(Run_Grid.Columns("add_time").Value) Then
                                Speak_Error ("必须先输入 " + Run_Grid.Columns("add_time").Caption)
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                        End If
                         If IsEmpty(Run_Grid.Columns("cai_number").Value) Then
                                Speak_Error ("必须先输入 " + Run_Grid.Columns("cai_number").Caption)
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                        End If
                    If Check_Weight(Pei_fan_table_Input.Pei_Fan_Mathine, Pei_fan_table_Input.Pei_Fan_Number, _
                                        Run_Grid.Columns("add_time").Value, 1, Run_Grid.Columns("weight").Value, Run_Grid.Columns("cai_number").Value, _
                                        Run_Grid.Columns("sort_idx").Value) <> 0 Then
                                Run_Grid.SetFocus
                                Cancel = True
                    End If
                Case "add_time"
                    If Not Check_Float(Run_Grid.Columns("add_time").Value) Then
                        Cancel = True
                        Run_Grid.SetFocus
                        Exit Sub
                    End If
                    If Run_Grid.Columns("add_time").Value <> 1 And Run_Grid.Columns("add_time").Value <> 2 Then
                        Speak_Error ("加料次数 的范围是<1 或>2 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
            End Select
            
                
End Sub



Private Sub run_Grid_BeforeUpdate(Cancel As Integer)
        Dim Temp_weight, Temp_fast, Temp_drop, Temp_gon As Single
        Dim i%
        If change_flag = UNCHANGE Then
                    Exit Sub
        End If
    For i% = 0 To Run_Grid.Columns.count - 1
            If Run_Grid.Columns(i%).Text = "" And Run_Grid.Columns(i%).Visible Then
                Call Speak_Error("应输入" & Run_Grid.Columns(i%).Caption)
                Run_Grid.SetFocus
                Cancel = True
                Exit Sub
           End If
           If Run_Grid.Columns(i%).Text <> "" And Run_Grid.Columns(i%).Text = "0" And Run_Grid.Columns(i%).Visible Then
                Call Speak_Error("应输入" & Run_Grid.Columns(i%).Caption & ">0")
                
                Cancel = True
                Run_Grid.SetFocus
                Exit Sub
           
           End If
    Next i%
    
        
            If Run_Grid.Columns("weight").DataChanged Or Run_Grid.Columns("gon_cha").DataChanged Or Run_Grid.Columns("fast_do").DataChanged Or Run_Grid.Columns("drop_do").DataChanged Then
                    Temp_weight = CInt(Run_Grid.Columns("weight").Text)
                    Temp_fast = CSng(vFieldVal(Run_Grid.Columns("fast_do").Text))
                    Temp_drop = CSng(vFieldVal(Run_Grid.Columns("drop_do").Text))
                    Temp_gon = CSng(vFieldVal(Run_Grid.Columns("gon_cha").Text))
                    
                    If Temp_weight < Temp_fast Then
                                Call Speak_Error("重量错误，应重量>=快慢提前量")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                    End If
                    If Temp_fast < Temp_drop Then
                                Call Speak_Error("重量错误，应快慢提前量>=点动提前量")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                    End If
                    If Temp_drop < Temp_gon Then
                                Call Speak_Error("公差错误，应公差<=点动量")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                    End If
        End If
End Sub

Private Sub Run_Grid_Error(ByVal DataError As Integer, Response As Integer)
            Response = 0
End Sub

Private Sub run_Grid_GotFocus()
        If Run_Grid.col = 0 Then
            Run_Grid.col = 2
        End If
End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, Flag As Integer
    Select Case KeyCode
        Case vbKeyUp
                        Run_Grid.AllowAddNew = False
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
              If Run_Grid.VisibleRows = 1 Then
                    Exit Sub
              End If
        Select Case Run_Grid.col
                Case 0
                            If LastCol > 0 Then
                                Run_Grid.col = LastCol
                            Else
                                Run_Grid.col = 2
                            End If
                 Case 3
                    If LastCol = 4 Then
                            Run_Grid.col = 2
                    Else
                            Run_Grid.col = 4
                    End If
            End Select
            If Run_Grid.Columns("pei_number").Value = "" Then
                    old_index = DTan_hei_table.Recordset.RecordCount + 1
                    Run_Grid.Columns("Sort_Idx").Value = old_index
                    Run_Grid.Columns("pei_number").Value = Pei_fan_table_Input.Pei_Fan_Number
                    Run_Grid.Columns("add_time").Value = 1
            End If
            If Run_Grid.Columns("Sort_idx").Value = "" Then
                    old_index = DTan_hei_table.Recordset.AbsolutePosition + 1
                    Run_Grid.Columns("Sort_Idx").Value = old_index
            End If


            Exit Sub
errorHandle:
        Speak_Error ("须选择配方")
End Sub

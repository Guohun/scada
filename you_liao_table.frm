VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form You_liao_table_Input 
   Caption         =   "You_liao"
   ClientHeight    =   5616
   ClientLeft      =   -72
   ClientTop       =   432
   ClientWidth     =   9504
   KeyPreview      =   -1  'True
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5616
   ScaleWidth      =   9504
   Begin VB.Data Cai_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cai_liao_table"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.Data DYou_liao_table 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\ljxt.Mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "you_liao_table"
      Top             =   4320
      Visible         =   0   'False
      Width           =   912
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "you_liao_table.frx":0000
      Height          =   1572
      Left            =   120
      OleObjectBlob   =   "you_liao_table.frx":0013
      TabIndex        =   1
      Top             =   3360
      Width           =   9132
   End
   Begin MSDBGrid.DBGrid run_Grid 
      Bindings        =   "you_liao_table.frx":105D
      Height          =   2532
      Left            =   120
      OleObjectBlob   =   "you_liao_table.frx":1077
      TabIndex        =   0
      Top             =   720
      Width           =   9132
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   7560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -120
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
      Picture         =   "you_liao_table.frx":2FC6
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   4560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -120
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
      Picture         =   "you_liao_table.frx":3418
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   1440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -120
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
      Picture         =   "you_liao_table.frx":386A
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   732
      Left            =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   -120
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
      MouseIcon       =   "you_liao_table.frx":3CBC
      Picture         =   "you_liao_table.frx":410E
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   3000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   -120
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
      MouseIcon       =   "you_liao_table.frx":4560
      Picture         =   "you_liao_table.frx":49B2
   End
   Begin VB.Label list_name 
      Caption         =   "list_name"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "You_liao_table_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msTBookMark As String
Dim You_ban_weight   As Single
Dim You_weight(0 To 20)   As Single
Dim You_ban_Name     As String
Dim change_flag      As Boolean

'检测油重量
Public Sub Add_list()
    Dim number_str    As String * 8, name_str As String * 15
     Dim Total_Weight As Single
     Dim Wna   As Integer
     Dim i%
     i = 0
    Total_Weight = 0
     DYou_liao_table.Refresh
     If DYou_liao_table.Recordset.RecordCount > 0 Then
         Do While Not DYou_liao_table.Recordset.EOF
                        If Not IsNull(DYou_liao_table.Recordset.Fields("weight").Value) Then
                            You_weight(i) = DYou_liao_table.Recordset.Fields("weight").Value
                        Else
                            You_weight(i) = 0
                        End If
                        Total_Weight = Total_Weight + You_weight(i)
                        DYou_liao_table.Recordset.MoveNext
                        i = i + 1
        Loop
        DYou_liao_table.Recordset.MoveFirst
    End If
    If Total_Weight > You_ban_weight Then
                Speak_Error (You_ban_Name & "超重: '" & Total_Weight - You_ban_weight & "'公斤")
    End If
    
End Sub










Private Sub CmdFind_Click()
 'On Error GoTo FindErr
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
  
   If DYou_liao_table.Recordset.Fields(Run_Grid.Columns(gsFindFiel).DataField).Type = dbText Then
        sTmp = "[" + Run_Grid.Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + "'" & Trim(gsFindExpr) & "'"
    Else
        sTmp = "[" + Run_Grid.Columns(gsFindFiel).DataField + "]" + space(1) + gsFindOp + space(1) + Trim(gsFindExpr)
   End If
  i = frmFind.lstFields.ListIndex
  sBookMark = DYou_liao_table.Recordset.Bookmark
  'search for the record
  Select Case gsFindType
    Case 0
      DYou_liao_table.Recordset.FindFirst sTmp
    Case 1
      DYou_liao_table.Recordset.FindNext sTmp
    Case 2
      DYou_liao_table.Recordset.FindPrevious sTmp
    Case 3
      DYou_liao_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = DYou_liao_table.Recordset.NoMatch
  If mbNotFound = True Then MsgBox "没发现!"
AfterWhile:
  Screen.MousePointer = vbDefault

  If gbFindFaile = True Then   'go back to original row
    DYou_liao_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "没发现", 48
    DYou_liao_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = DYou_liao_table.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    DYou_liao_table.Recordset.Bookmark = sBookMark
  End If

  mbDataChanged = False

  Exit Sub

FindErr:
    Exit Sub
    mbNotFound = True
End Sub










Private Sub DBGrid1_DblClick()
                    Run_Grid.Columns("cai_number").Text = DBGrid1.Columns("cai_number").Text
                    Run_Grid.Columns("name").Text = DBGrid1.Columns("name").Text
                    Run_Grid.Columns("dou").Text = DBGrid1.Columns("dou").Text
                    Run_Grid.Columns("wna").Text = DBGrid1.Columns("wna").Text
                    Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                    change_flag = CHANGE
                    DBGrid1.SetFocus

End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
                   Dim i%
           If KeyCode = vbKeyTab Then
            'run_Grid.SetFocus
                Exit Sub
        End If
     If KeyCode = 13 Then
                    change_flag = CHANGE
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
Call Add_Win(Hwnd)
 Width = MAIN_MDI.Width  ' Set width of form.
 Height = MAIN_MDI.Height - 600 ' Set height of form.
 Left = 0
 Top = 0
 change_flag = UNCHANGE
 DYou_liao_table.DatabaseName = Data_Path & "\ljxt.mdb"
 Cai_Data.DatabaseName = Data_Path & "\ljxt.mdb"
 If Pei_fan_table_Input.You_Liao_Flag = 1 Then
    DYou_liao_table.RecordSource = "select * from You_liao_table where wna='2' and pei_number like '" & Pei_fan_table_Input.Pei_Fan_Number & "' and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & "  order by sort_idx "
    Cai_Data.RecordSource = "select * from cai_liao_table  where  wna=2  and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & " and  cai_number <> '  '   " & " order by dou"
  Else
    DYou_liao_table.RecordSource = "select * from You_liao_table where wna='3' and pei_number like '" & Pei_fan_table_Input.Pei_Fan_Number & "' and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & "  order by sort_idx "
    Cai_Data.RecordSource = "select * from cai_liao_table  where  wna=3  and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & " and  cai_number <> '  '   " & " order by dou"
End If
 DYou_liao_table.Refresh
 Cai_Data.Refresh
  DBGrid1.Refresh
   If FirstDo() = 0 Then
                Exit Sub
    End If
 Run_Grid.Caption = Pei_fan_table_Input.Pei_Fan_Number & "油料配方"


 If DYou_liao_table.Recordset.RecordCount = 0 Then
     Run_Grid.Columns("pei_number").Value = Pei_fan_table_Input.Pei_Fan_Number             'index
     Run_Grid.Columns("add_time").Text = 1
     Run_Grid.Columns("Sort_idx").Value = 1

 End If
'Call Add_list
End Sub

Private Sub Form_Paint()
        If China_English = CHINA Then
          Caption = "油料录入"

          cmdDelete.Caption = "F10删除记录"
          CmdFind.Caption = "F6查询记录"
          Exit_Command.Caption = "ESC退出"
          Print_Command.Caption = "F4打印"

          Insert_Command.Caption = "F3插入"
          
          Run_Grid.Columns("sort_idx").Caption = "目录"

          Run_Grid.Columns("pei_number").Caption = "配方编号"
          Run_Grid.Columns("cai_number").Caption = "材料编号"
          Run_Grid.Columns("name").Caption = "材料名称"
          Run_Grid.Columns("dou").Caption = "斗    号"
          Run_Grid.Columns("weight").Caption = "所需重量"
          Run_Grid.Columns("fast_do").Caption = "快慢提前量"
          Run_Grid.Columns("drop_do").Caption = "点动提前量:"
          Run_Grid.Columns("gon_cha").Caption = "允许公差"
          Run_Grid.Columns("control_time").Caption = "控制时间"
          Run_Grid.Columns("add_time").Caption = "加料次数"
          Run_Grid.Columns("stop_time").Caption = "稳定时间"
          Run_Grid.Columns("wna").Caption = "秤"

'
           DBGrid1.Columns("cai_number").Caption = "材料编号"
           DBGrid1.Columns("name").Caption = "材料名称"
           DBGrid1.Columns("wna").Caption = "秤名"
           DBGrid1.Columns("dou").Caption = "斗名"
           DBGrid1.Columns("jiao").Caption = "胶状"
           DBGrid1.Columns("mathine").Caption = "机组"
           list_name.Caption = "用TAB键盘切换 表格"
        Else
          Caption = "you"

          cmdDelete.Caption = "F10 del"
          CmdFind.Caption = "F6 search"
          Exit_Command.Caption = "ESC "
          Print_Command.Caption = "F4 prin"

          Insert_Command.Caption = "F3 ins"
          
          Run_Grid.Columns("sort_idx").Caption = " index"

          Run_Grid.Columns("pei_number").Caption = "A"
          Run_Grid.Columns("cai_number").Caption = "B"
          Run_Grid.Columns("name").Caption = "C"
          Run_Grid.Columns("dou").Caption = "D"
          Run_Grid.Columns("weight").Caption = "D"
          Run_Grid.Columns("fast_do").Caption = "W"
          Run_Grid.Columns("drop_do").Caption = "F:"
          Run_Grid.Columns("gon_cha").Caption = "G"
          Run_Grid.Columns("control_time").Caption = "K"
          Run_Grid.Columns("add_time").Caption = "D"
          Run_Grid.Columns("stop_time").Caption = "W"
          Run_Grid.Columns("wna").Caption = "C"

'
           DBGrid1.Columns("cai_number").Caption = "B"
           DBGrid1.Columns("name").Caption = "C"
           DBGrid1.Columns("wna").Caption = "Z"
           DBGrid1.Columns("dou").Caption = "D"
           DBGrid1.Columns("jiao").Caption = "J"
           DBGrid1.Columns("mathine").Caption = "M"
           list_name.Caption = " "
        End If
End Sub



Private Function FirstDo() As Integer
        Dim i, count As Integer
        Dim gsDBName As Database
        Dim gsThisSet As Recordset

        You_ban_weight = 0
        
        
        Set gsDBName = OpenDatabase("c:\ljxt\COMM.mdb")
        Set gsThisSet = gsDBName.OpenRecordset( _
            "select code ,name,max  from  send_table   Where mathine = " & Pei_fan_table_Input.Pei_Fan_Mathine, dbOpenDynaset)
            Do While Not gsThisSet.EOF
                If Not IsEmpty(gsThisSet.Fields(0).Value) Then
                        If gsThisSet.Fields(0).Value = 2 Then
                            You_ban_weight = gsThisSet.Fields(2).Value
                            You_ban_Name = gsThisSet.Fields(1).Value
                        End If
                End If
                gsThisSet.MoveNext
             Loop
        gsThisSet.Close
        gsDBName.Close
        
        If You_ban_weight = 0 Then
                FirstDo = 0
                Call Speak_Error("油秤没定义或重量为零")
                Exit Function
        End If
        
        FirstDo = 1

        For i = 0 To Run_Grid.Columns.count - 1
               Run_Grid.Columns(i).Width = 700
        Next i
        Run_Grid.Columns("password").Visible = False
        Run_Grid.Columns("name").Width = 1700
        Run_Grid.Columns("pei_number").Visible = False
        Run_Grid.Columns("wna").Visible = False
        Run_Grid.Columns("mathine").Visible = False
        
        Run_Grid.Font.Size = 10
        Run_Grid.HeadFont.Size = 12
        Run_Grid.RowHeight = 400
        Run_Grid.HeadLines = 2
                
        DBGrid1.Columns("jiao").Visible = False
        DBGrid1.Font.Size = 10
        DBGrid1.HeadFont.Size = 10
        DBGrid1.RowHeight = 300
        DBGrid1.HeadLines = 2
        
End Function


Private Sub Form_Unload(Cancel As Integer)
        Dim strsqlrestore   As String
        strsqlrestore = "delete  from   you_liao_table    where ([cai_number] is null or weight is null) "
        Debug.Print strsqlrestore
        DYou_liao_table.Database.Execute strsqlrestore, dbFailOnError
     Call Del_Win(Hwnd)
End Sub

Private Sub cmdDelete_Click()
    Dim Temp_record As Integer
    If DYou_liao_table.Recordset.RecordCount <= 0 Then
                    Exit Sub
    End If
    If DYou_liao_table.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择配方")
        Exit Sub
   End If
If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
          Temp_record = DYou_liao_table.Recordset.Fields("sort_idx").Value
          Run_Grid.AllowAddNew = False
           DYou_liao_table.Recordset.Delete
           i% = 1
           DYou_liao_table.Recordset.MoveFirst
            Do While Not DYou_liao_table.Recordset.EOF()
                   If IsNull(DYou_liao_table.Recordset.Fields("weight").Value) Then
                       You_weight(i% - 1) = 0
                    Else
                            You_weight(i% - 1) = DYou_liao_table.Recordset.Fields("weight").Value
                    End If

                
                  DYou_liao_table.Recordset.Edit
                  DYou_liao_table.Recordset.Fields("sort_idx").Value = i%
                  DYou_liao_table.Recordset.Update
                  i% = i% + 1
                  DYou_liao_table.Recordset.MoveNext
            Loop
            With DYou_liao_table.Recordset
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








Private Sub Insert_Command_Click()
 Dim temp%
 'On Error GoTo ErrorHandler:
  If Run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能插入")
            Run_Grid.SetFocus
            Exit Sub
 End If

  If DYou_liao_table.Recordset.RecordCount > 0 Then
        If DYou_liao_table.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
        Exit Sub
      End If
      temp% = Run_Grid.Columns(0).Value
'        run_Grid.AllowAddNew = True
        i% = temp%
        DYou_liao_table.Recordset.FindFirst "sort_idx=" & i%
        Do While Not DYou_liao_table.Recordset.NoMatch
                DYou_liao_table.Recordset.Edit
                DYou_liao_table.Recordset.Fields("sort_idx").Value = i% + 1
                DYou_liao_table.Recordset.Update
                i% = i% + 1
                DYou_liao_table.Recordset.FindNext "sort_idx=" & i%
        Loop
        
        Dim strsqlrestore   As String
         strsqlrestore = "insert  into you_liao_table (sort_idx,pei_number,add_time)  " _
                                          & "  values (" & temp% & ",'" & Pei_fan_table_Input.Pei_Fan_Number & "'," & 1 & ")"
                                          
        DYou_liao_table.Database.Execute strsqlrestore, dbFailOnError
        
        DYou_liao_table.Refresh
        DYou_liao_table.Recordset.FindFirst "sort_idx=" & temp%
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
  With DYou_liao_table
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
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    'Printer.FontName = "Courier"
    'Printer.FontSize = 12
    
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
    
    DYou_liao_table.Recordset.MoveFirst
    j% = 0
    Do While Not DYou_liao_table.Recordset.EOF()
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
         DYou_liao_table.Recordset.MoveNext
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
'    Printer.Print Space(10) + "总数:" & CStr(DYou_liao_table.Recordset.RecordCount)
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







Private Sub run_Grid_AfterUpdate()
       change_flag = UNCHANGE
End Sub

Private Sub Run_Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
            Select Case Run_Grid.Columns(ColIndex).DataField
                Case "weight", "fast_do", "drop_do", "gon_cha", "control_time"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) And KeyAscii <> Asc(".") Then
                        Speak_Error ("必须输入重量为数值 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                    change_flag = CHANGE
                Case "add_time", "dou", "stop_time"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入为数字 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                    change_flag = CHANGE
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
                                    Cancel = True
                                    Run_Grid.SetFocus
                                Exit Sub
                   End If
                            Call DBGrid1_DblClick
                            Run_Grid.SetFocus
                End If
                Case "dou"
                        If Not (Check_Float(Run_Grid.Columns("dou").Value)) Then
                                     Speak_Error ("必须输入正确斗号")
                                      Cancel = True
                                      Run_Grid.SetFocus
                                      Exit Sub
                        End If
                        If Run_Grid.Columns("dou").Value <> OldValue Then
                                Cai_Data.Recordset.FindFirst "dou=  " & Run_Grid.Columns("dou").Value
                                If Cai_Data.Recordset.NoMatch Then
                                    Speak_Error ("无此斗号 ")
                                    Cancel = True
                                    Run_Grid.SetFocus
                                Exit Sub
                               End If
                            Call DBGrid1_DblClick
                            Run_Grid.SetFocus
                    End If
                Case "gon_cha", "stop_time", "control_time", "drop_do", "fast_do"
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
                        If Pei_fan_table_Input.You_Liao_Flag = 1 Then
                        If Check_Weight(Pei_fan_table_Input.Pei_Fan_Mathine, Pei_fan_table_Input.Pei_Fan_Number, _
                                        Run_Grid.Columns("add_time").Value, 2, Run_Grid.Columns("weight").Value, Run_Grid.Columns("cai_number").Value, _
                                        Run_Grid.Columns("sort_idx").Value) <> 0 Then
                                Run_Grid.SetFocus
                                Cancel = True
                        End If
                    Else
                        If Check_Weight(Pei_fan_table_Input.Pei_Fan_Mathine, Pei_fan_table_Input.Pei_Fan_Number, _
                                        Run_Grid.Columns("add_time").Value, 2, Run_Grid.Columns("weight").Value, Run_Grid.Columns("cai_number").Value, _
                                        Run_Grid.Columns("sort_idx").Value) <> 0 Then
                                Run_Grid.SetFocus
                                Cancel = True
                        End If
                    End If
                    Exit Sub
                    
                    Total_Weight = 0
                    If IsNull(Run_Grid.Columns("weight").Value) Then
                        You_weight(Run_Grid.Columns("sort_idx").Value - 1) = 0
                    Else
                        You_weight(Run_Grid.Columns("sort_idx").Value - 1) = Run_Grid.Columns("weight").Value
                    End If

                    Debug.Print Run_Grid.Columns("weight").Value
                    Dim temp As Single
                    If DYou_liao_table.Recordset.RecordCount > Run_Grid.VisibleRows Then
                            temp = DYou_liao_table.Recordset.RecordCount
                    Else
                            temp = Run_Grid.VisibleRows
                    End If
                    For i% = 0 To (temp - 1)
                            Total_Weight = Total_Weight + You_weight(i%)
                            Debug.Print You_weight(i%)
                    Next i%
                    Debug.Print Total_Weight
                    If Total_Weight > You_ban_weight Then
                          Speak_Error ("必须输入秤重小于" & You_ban_weight)
                          Cancel = True
                           If OldValue <> "" Then
                                You_weight(Run_Grid.Columns("sort_idx").Value - 1) = CSng(OldValue)
                            Else
                                You_weight(Run_Grid.Columns("sort_idx").Value - 1) = 0
                            End If
                            Run_Grid.SetFocus
                          Exit Sub
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
                        Exit Sub
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
           If Run_Grid.Columns(i%).Text <> "" And Run_Grid.Columns(i%).Text = "0" Then
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
              
              If Run_Grid.VisibleRows < 1 Then
                    Exit Sub
              End If
              If Run_Grid.col = 0 And LastCol > 0 Then
                    Run_Grid.col = LastCol
              End If
             If Run_Grid.col = 3 Then
                    If LastCol = 4 Then
                            Run_Grid.col = 2
                    Else
                        Run_Grid.col = 4
                    End If
             End If
            If Run_Grid.Columns("pei_number").Value = "" Then
                    old_index = DYou_liao_table.Recordset.RecordCount + 1
                    Run_Grid.Columns("Sort_Idx").Value = old_index
                    Run_Grid.Columns("pei_number").Value = Pei_fan_table_Input.Pei_Fan_Number
                    Run_Grid.Columns("add_time").Value = 1
            End If
            If Run_Grid.Columns("Sort_idx").Value = "" Then
                    old_index = DYou_liao_table.Recordset.AbsolutePosition + 1
                    Run_Grid.Columns("Sort_Idx").Value = old_index
            End If
            Exit Sub
errorHandle:
        Speak_Error ("须选择配方")
End Sub

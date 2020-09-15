VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form jiao_liao_table_Input 
   Caption         =   "jiao_liao_table"
   ClientHeight    =   5580
   ClientLeft      =   396
   ClientTop       =   1452
   ClientWidth     =   8772
   KeyPreview      =   -1  'True
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5580
   ScaleWidth      =   8772
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "jiao_liao_table.frx":0000
      Height          =   1620
      Left            =   216
      OleObjectBlob   =   "jiao_liao_table.frx":0013
      TabIndex        =   1
      Top             =   3696
      Width           =   7332
   End
   Begin MSDBGrid.DBGrid run_Grid 
      Bindings        =   "jiao_liao_table.frx":105D
      Height          =   2700
      Left            =   216
      OleObjectBlob   =   "jiao_liao_table.frx":1078
      TabIndex        =   0
      Top             =   792
      Width           =   7332
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   6696
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   24
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
      Font3D          =   2
      Picture         =   "jiao_liao_table.frx":2C7B
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   3696
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   24
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
      Picture         =   "jiao_liao_table.frx":30CD
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   576
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   24
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
      Picture         =   "jiao_liao_table.frx":351F
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   732
      Left            =   5256
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   24
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
      MouseIcon       =   "jiao_liao_table.frx":3971
      Picture         =   "jiao_liao_table.frx":3DC3
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   2136
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   24
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
      MouseIcon       =   "jiao_liao_table.frx":4215
      Picture         =   "jiao_liao_table.frx":4667
   End
   Begin VB.Data Cai_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cai_liao_table"
      Top             =   4344
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.Data Djiao_liao_table 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\ljxt.Mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   336
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "jiao_liao_table"
      Top             =   4200
      Visible         =   0   'False
      Width           =   912
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
      Height          =   4452
      Left            =   7800
      TabIndex        =   7
      Top             =   840
      Width           =   492
   End
End
Attribute VB_Name = "jiao_liao_table_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msTBookMark As String
Dim Jiao_Ban_weight As Single
Dim change_flag As Boolean

Dim Tan_Weight(0 To 20) As Single


Public Sub Add_list()
    Dim number_str    As String * 8, name_str As String * 15
     Dim Total_Weight As Single
     Dim i%
    Total_Weight = 0
    i = 0
     Tan_Weight(i) = 0
     Djiao_liao_table.Refresh
     If Djiao_liao_table.Recordset.RecordCount > 0 Then
         Djiao_liao_table.Recordset.MoveFirst
         Do While Not Djiao_liao_table.Recordset.EOF
                        name_str = vFieldVal(Djiao_liao_table.Recordset.Fields(3).Value)
                        If Not IsNull(Djiao_liao_table.Recordset.Fields("weight").Value) Then
                            Tan_Weight(i) = Djiao_liao_table.Recordset.Fields("weight").Value
                            Total_Weight = Total_Weight + Tan_Weight(i)
                            i = i + 1
                            Tan_Weight(i) = 0
                        End If
                        Djiao_liao_table.Recordset.MoveNext
        Loop
        Djiao_liao_table.Recordset.MoveFirst
    End If
    If Total_Weight > Jiao_Ban_weight Then
                Speak_Error ("超重: '" & Total_Weight - Tan_Ban_weight & "'公斤")
    End If
End Sub








Private Sub cmdDelete_Click()
    If Djiao_liao_table.Recordset.RecordCount <= 0 Then
                    Exit Sub
    End If
  If Djiao_liao_table.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择配方")
            Run_Grid.SetFocus
        Exit Sub
   End If
If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
         Dim Temp_record%
         Temp_record = Djiao_liao_table.Recordset.Fields("sort_idx").Value
           Djiao_liao_table.Recordset.Delete
           i% = 1
           Djiao_liao_table.Recordset.MoveFirst
            Do While Not Djiao_liao_table.Recordset.EOF()
                If IsNull(Djiao_liao_table.Recordset.Fields("weight").Value) Then
                    Tan_Weight(i - 1) = 0
                Else
                    Tan_Weight(i - 1) = Djiao_liao_table.Recordset.Fields("weight").Value
                End If
                  Djiao_liao_table.Recordset.Edit
                  Djiao_liao_table.Recordset.Fields("sort_idx").Value = i%
                  Djiao_liao_table.Recordset.Update
                  i% = i% + 1
                  Djiao_liao_table.Recordset.MoveNext
            Loop
            With Djiao_liao_table.Recordset
                .FindFirst "sort_idx=" & Temp_record
                If .NoMatch Then
                    If .RecordCount > 0 Then
                     .MoveLast
                    End If
                End If
            End With
            
            Run_Grid.SetFocus
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then Djiao_liao_table.Recordset.Delete
End If
End Sub


Private Sub CmdFind_Click()
 On Error GoTo FindErr
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
FindStart:

   frmFind.lstFields.AddItem Run_Grid.Columns("cai_number").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("name").Caption
   frmFind.lstFields.AddItem Run_Grid.Columns("weight").Caption

      
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
  sBookMark = Djiao_liao_table.Recordset.Bookmark
  'search for the record
  Select Case gnFindType
    Case 0
      Djiao_liao_table.Recordset.FindFirst sTmp
    Case 1
      Djiao_liao_table.Recordset.FindNext sTmp
    Case 2
      Djiao_liao_table.Recordset.FindPrevious sTmp
    Case 3
      Djiao_liao_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = Djiao_liao_table.Recordset.NoMatch

AfterWhile:

  Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    Djiao_liao_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "没发现", 48
    Djiao_liao_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = Djiao_liao_table.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    Djiao_liao_table.Recordset.Bookmark = sBookMark
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
                    Run_Grid.Columns("st").Text = DBGrid1.Columns("jiao").Text
                    Run_Grid.Columns("wna").Text = DBGrid1.Columns("wna").Text
                    Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                    
'                    i% = DTan_hei_table.Recordset.AbsolutePosition
'                  Call Add_list
 '                  DTan_hei_table.Recordset.AbsolutePosition = i%
                   DBGrid1.SetFocus

End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
               Dim i%
     If KeyCode = vbKeyTab Then
        Exit Sub
     End If
     If KeyCode = 13 Then
                    Run_Grid.Columns("cai_number").Text = DBGrid1.Columns("cai_number").Text
                    Run_Grid.Columns("name").Text = DBGrid1.Columns("name").Text
                    Run_Grid.Columns("st").Text = DBGrid1.Columns("jiao").Text
                    Run_Grid.Columns("wna").Text = DBGrid1.Columns("wna").Text
                    Run_Grid.Columns("mathine").Text = DBGrid1.Columns("mathine").Text
                   DBGrid1.SetFocus
          End If
 If (KeyCode <> 40 And KeyCode <> 38) Then KeyCode = 0
End Sub





Private Sub Exit_Command_Click()
        Unload Me
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
'On Error Resume Next
     Call Add_Win(Hwnd)
     If MAIN_MDI.Width > 10 Then
                Width = MAIN_MDI.Width - 10 ' Set width of form.
    End If
      If MAIN_MDI.Height > 100 Then
            Height = MAIN_MDI.Height - 100 ' Set height of form.
      End If
     Left = 0
     Top = 0
     'Me.WindowState = 2
    
     
     change_flag = UNCHANGE
  Cai_Data.DatabaseName = Data_Path & "\ljxt.mdb"
 Djiao_liao_table.DatabaseName = Data_Path & "\ljxt.mdb"

   Djiao_liao_table.RecordSource = "select * from jiao_liao_table where pei_number like '" & Pei_fan_table_Input.Pei_Fan_Number & "'  and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & "  order by sort_idx "
   Djiao_liao_table.Refresh
   
    Cai_Data.RecordSource = "select * from cai_liao_table  where  wna=4 and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & "  and  cai_number <> '  '   " & " order by jiao"
 Cai_Data.Refresh
  DBGrid1.Refresh

 Call FirstDo

 If Djiao_liao_table.Recordset.RecordCount = 0 Then
     
      Run_Grid.Columns("pei_number").Text = Pei_fan_table_Input.Pei_Fan_Number
      Run_Grid.Columns("sort_idx").Text = 1
     
 End If
  'Call Add_list
End Sub

Private Function FirstDo() As Integer
Dim i, count As Integer
Dim gsDBName As Database
Dim gsThisSet As Recordset
       First_Do = 1
       
        Set gsDBName = OpenDatabase("c:\ljxt\COMM.mdb")
        Set gsThisSet = gsDBName.OpenRecordset( _
           "select code,name,max  from  send_table   Where mathine = " & Pei_fan_table_Input.Pei_Fan_Mathine, dbOpenDynaset)
            Do While Not gsThisSet.EOF
                If Not IsEmpty(gsThisSet.Fields(0).Value) Then
                        If gsThisSet.Fields(0).Value = 4 Then
                            Jiao_Ban_weight = gsThisSet.Fields(2).Value
                        Exit Do
                        End If
                End If
                gsThisSet.MoveNext
             Loop
        gsThisSet.Close
        gsDBName.Close
        
        For i = 0 To Run_Grid.Columns.count - 1
               Run_Grid.Columns(i).Width = 800
        Next i
        Run_Grid.Columns("password").Visible = False
        Run_Grid.Columns("name").Width = 1800
        Run_Grid.Columns("pei_number").Visible = False
        Run_Grid.Columns("CONTROL_TIME").Visible = False
        Run_Grid.Columns("STOP_TIME").Visible = False
        Run_Grid.Columns("mathine").Visible = False
        Run_Grid.Columns("wna").Visible = False
        
        Run_Grid.Columns("ADD_TIME").Visible = True
        
        Run_Grid.Font.Size = 10
        Run_Grid.HeadFont.Size = 12
        Run_Grid.RowHeight = 400
        Run_Grid.HeadLines = 2
        DBGrid1.Columns("dou").Visible = False
        DBGrid1.Columns("jiao").Visible = True
        DBGrid1.Font.Size = 10
        DBGrid1.HeadFont.Size = 10
        DBGrid1.RowHeight = 300
        DBGrid1.HeadLines = 2


End Function

Private Sub Form_Paint()
If China_English = CHINA Then
    Caption = "胶"
  cmdDelete.Caption = "F10删除记录"
  CmdFind.Caption = "F6查询记录"
  Exit_Command.Caption = "ESC退出"
  Print_Command.Caption = "F4打印"
  Insert_Command.Caption = "F3插入"
         
  Run_Grid.Columns("sort_idx").Caption = "目录"
  Run_Grid.Caption = Pei_fan_table_Input.Pei_Fan_Number & "胶料配方"
'  Frame2.Caption = "移动记录"
  Run_Grid.Columns("pei_number").Caption = "配方编号:"
  Run_Grid.Columns("cai_number").Caption = "材料编号:"
  Run_Grid.Columns("name").Caption = "材料名称:"
  Run_Grid.Columns("wna").Caption = "秤名："
  Run_Grid.Columns("st").Caption = "胶状:"
  Run_Grid.Columns("gon_cha").Caption = "公差"
  Run_Grid.Columns("add_time").Caption = "加料次数"
  Run_Grid.Columns("weight").Caption = "所需重量:"

  list_name.Caption = "用TAB 键切换窗口"
  
  
           DBGrid1.Columns("cai_number").Caption = "材料编号"
           DBGrid1.Columns("name").Caption = "材料名称"
           DBGrid1.Columns("wna").Caption = "秤名"
           DBGrid1.Columns("dou").Caption = "斗名"
           DBGrid1.Columns("jiao").Caption = "胶状"
           DBGrid1.Columns("mathine").Caption = "机组"
  
'  Cai_List_Name.Caption = "已有材料"
Else
    Caption = "jiao"
  cmdDelete.Caption = "F10 Delete"
  CmdFind.Caption = "F6 ind"
  Exit_Command.Caption = "ESC"
  Print_Command.Caption = "F4print"
  Insert_Command.Caption = "F3 insert"
         
  Run_Grid.Columns("sort_idx").Caption = "index"
  Run_Grid.Caption = ""
  Run_Grid.Columns("pei_number").Caption = "H:"
  Run_Grid.Columns("cai_number").Caption = "V:"
  Run_Grid.Columns("name").Caption = "N:"
  Run_Grid.Columns("wna").Caption = "V："
  Run_Grid.Columns("st").Caption = "H:"
  Run_Grid.Columns("gon_cha").Caption = "G"
  Run_Grid.Columns("add_time").Caption = "J"
  Run_Grid.Columns("weight").Caption = "G:"

  list_name.Caption = "TAB"
  
  
           DBGrid1.Columns("cai_number").Caption = "C"
           DBGrid1.Columns("name").Caption = "M"
           DBGrid1.Columns("wna").Caption = "B"
           DBGrid1.Columns("dou").Caption = "D"
           DBGrid1.Columns("jiao").Caption = "H"
           DBGrid1.Columns("mathine").Caption = "J"
  
  list_name.Caption = "use Tab key "
  
End If
End Sub

Private Sub Form_Terminate()
Call Del_Win(Hwnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
       Dim strsqlrestore   As String
       strsqlrestore = "delete  from   jiao_liao_table    where [cai_number] is null"
       Djiao_liao_table.Database.Execute strsqlrestore, dbFailOnError
       
   Call Del_Win(Hwnd)

End Sub







Private Sub Print_Command_Click()
   'On Error GoTo error_doing

   Dim Y  As Long
   Dim i%, k%
   Dim temp(0 To 20) As Single

   Dim book   As Variant
   book = Run_Grid.Bookmark
   
   Screen.MousePointer = vbHourglass
     If Printer.ScaleWidth < Run_Grid.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        Run_Grid.SetFocus
        Exit Sub
    End If
  With Djiao_liao_table
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
    'Printer.FontName = "Courier"
   ' Printer.FontSize = 20

        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
        Printer.Print Run_Grid.Caption + space(4) + Date$
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
    If Djiao_liao_table.Recordset.RecordCount > 0 Then
            Djiao_liao_table.Recordset.MoveFirst
    End If
    j% = 0
    Do While Not Djiao_liao_table.Recordset.EOF()
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
         Djiao_liao_table.Recordset.MoveNext
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
'    Printer.Print Space(10) + "总数:" & CStr(Djiao_liao_table.Recordset.RecordCount)
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
            change_flag = CHANGE
            Select Case Run_Grid.Columns(ColIndex).DataField
                Case "weight", "gon_cha"
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) And KeyAscii <> Asc(".") Then
                        Speak_Error ("必须输入重量为数值 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                Case "jiao", "add_time"
                    If (KeyAscii > Asc("2") Or KeyAscii < Asc("1")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入为数字1,2 ")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
            End Select

            
End Sub

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim i%, temp%
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
                Case "weight"
                       'If run_Grid.Columns(ColIndex).Value = "" Then
                        '             Speak_Error ("必须输入" + run_Grid.Columns(ColIndex).Caption)
                         '             Cancel = True
                          '            run_Grid.SetFocus
                           '           Exit Sub
                        'End If
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
                                        Run_Grid.Columns("add_time").Value, 4, Run_Grid.Columns("weight").Value, Run_Grid.Columns("cai_number").Value, _
                                        Run_Grid.Columns("sort_idx").Value) <> 0 Then
                                Run_Grid.SetFocus
                                Cancel = True
                    End If
                    Exit Sub
                    Total_Weight = 0
                    If IsNull(Run_Grid.Columns("weight").Value) Then
                        Tan_Weight(Run_Grid.Columns("sort_idx").Value - 1) = 0
                    Else
                        Tan_Weight(Run_Grid.Columns("sort_idx").Value - 1) = Run_Grid.Columns("weight").Value
                    End If
                    
                    If Djiao_liao_table.Recordset.RecordCount > Run_Grid.VisibleRows Then
                            temp = Djiao_liao_table.Recordset.RecordCount
                    Else
                            temp = Run_Grid.VisibleRows
                    End If
                    For i% = 0 To (temp - 1)
                            Total_Weight = Total_Weight + Tan_Weight(i%)
                            Debug.Print i%, Tan_Weight(i)
                    Next i%
                    Debug.Print Total_Weight
                    If Total_Weight > Jiao_Ban_weight Then
                          Speak_Error ("必须输入秤重小于" & Jiao_Ban_weight)
                          Cancel = True
                          If OldValue <> "" Then
                                Tan_Weight(Run_Grid.Columns("sort_idx").Value - 1) = CSng(OldValue)
                            Else
                                Tan_Weight(Run_Grid.Columns("sort_idx").Value - 1) = 0
                            End If
                            Run_Grid.SetFocus
                          Exit Sub
                    End If
                       Case "gon_cha"
                        'If run_Grid.Columns(ColIndex).Value = "" Then
                         '            Speak_Error ("必须输入" + run_Grid.Columns(ColIndex).Caption)
                          '            Cancel = True
                           '           run_Grid.SetFocus
                            '          Exit Sub
                        'End If
                         If Not Check_Float(Run_Grid.Columns(ColIndex).Value) Then
                                Speak_Error ("输入 " + Run_Grid.Columns(ColIndex).Caption + "非法")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                        End If
                Case "st"
                         If Not Check_Float(Run_Grid.Columns(ColIndex).Value) Then
                                Speak_Error ("输入 " + Run_Grid.Columns(ColIndex).Caption + "非法")
                                Cancel = True
                                Run_Grid.SetFocus
                                Exit Sub
                        End If
                
                Case Else
                   If Run_Grid.Columns(ColIndex).Text = "" And Run_Grid.Columns(ColIndex).Visible Then
                            Call Speak_Error("应输入" & Run_Grid.Columns(ColIndex).Caption)
                            Cancel = True
                            Run_Grid.SetFocus
                            Exit Sub
                        End If

                
End Select
    
    
End Sub


Private Sub run_Grid_BeforeUpdate(Cancel As Integer)
        Dim Temp_weight, Temp_gon As Single
        Dim i%
        If change_flag = UNCHANGE Then
                    Exit Sub
        End If
            For i% = 0 To Run_Grid.Columns.count - 1
                If Run_Grid.Columns(i%).Text = "" And Run_Grid.Columns(i%).Visible Then
                Call Speak_Error("应输入" & Run_Grid.Columns(i%).Caption)
                Cancel = True
                Run_Grid.SetFocus
                Exit Sub
            End If
        Next i%
                    Temp_weight = CInt(Run_Grid.Columns("weight").Text)
                    Temp_gon = CSng(vFieldVal(Run_Grid.Columns("gon_cha").Text))
                    If Temp_weight <= Temp_gon Then
                                Call Speak_Error("错误   应公差<" & Run_Grid.Columns("weight").Caption)
                                        Run_Grid.SetFocus
                                        Cancel = True
                                Exit Sub
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
                            Run_Grid.col = 1
                    End If
             Case 3, 4
                    If LastCol = 5 Then
                            Run_Grid.col = 2
                    Else
                          Run_Grid.col = 5
                    End If
             End Select
            
            If Run_Grid.Columns("pei_number").Value = "" Then
                    old_index = Djiao_liao_table.Recordset.RecordCount + 1
                    Run_Grid.Columns("Sort_Idx").Value = old_index
                    Run_Grid.Columns("pei_number").Value = Pei_fan_table_Input.Pei_Fan_Number
                    Run_Grid.Columns("st").Value = 1
            End If
            'If Run_Grid.Columns("Sort_idx").Value = "" Then
             '       old_index = Djiao_liao_table.Recordset.AbsolutePosition + 1
             '       Run_Grid.Columns("Sort_Idx").Value = old_index
           ' End If
            Exit Sub
errorHandle:
        Speak_Error ("Error : rowcol")
End Sub

Private Sub Insert_Command_Click()
Dim temp%
 'On Error GoTo ErrorHandler:
  If Run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能插入")
            Run_Grid.SetFocus
            Exit Sub
 End If

  If Djiao_liao_table.Recordset.RecordCount > 0 Then
        If Djiao_liao_table.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
        Exit Sub
      End If
      temp% = Run_Grid.Columns(0).Value
'        run_Grid.AllowAddNew = True
        i% = temp%
        Djiao_liao_table.Recordset.FindFirst "sort_idx=" & i%
        Do While Not Djiao_liao_table.Recordset.NoMatch
                Djiao_liao_table.Recordset.Edit
                Djiao_liao_table.Recordset.Fields("sort_idx").Value = i% + 1
                Djiao_liao_table.Recordset.Update
                i% = i% + 1
                Djiao_liao_table.Recordset.FindNext "sort_idx=" & i%
        Loop
        
        Dim strsqlrestore   As String
         strsqlrestore = "insert  into jiao_liao_table (sort_idx,pei_number)  " _
                                          & "  values (" & temp% & ",'" & Pei_fan_table_Input.Pei_Fan_Number & "'" & ")"
                                          
        Djiao_liao_table.Database.Execute strsqlrestore, dbFailOnError
        
        Djiao_liao_table.Refresh
        Djiao_liao_table.Recordset.FindFirst "sort_idx=" & temp%
        Run_Grid.Refresh
        
        Run_Grid.AllowAddNew = False
  End If
        Run_Grid.SetFocus
        
  Exit Sub
ErrorHandler:
    Speak_Error ("须选择配方")
    Run_Grid.SetFocus


End Sub

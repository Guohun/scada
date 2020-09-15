VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form lian_add_table_input 
   Caption         =   "lian_add_table"
   ClientHeight    =   5532
   ClientLeft      =   12
   ClientTop       =   1356
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5532
   ScaleWidth      =   9300
   Begin MSDBGrid.DBGrid run_Grid 
      Bindings        =   "lian_add_table.frx":0000
      Height          =   1092
      Left            =   240
      OleObjectBlob   =   "lian_add_table.frx":0018
      TabIndex        =   1
      Top             =   960
      Width           =   8892
   End
   Begin VB.Data lian_liao_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lain_liao_table"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid lian_DBGrid 
      Bindings        =   "lian_add_table.frx":1C53
      Height          =   3012
      Left            =   240
      OleObjectBlob   =   "lian_add_table.frx":1C6C
      TabIndex        =   0
      Top             =   2280
      Width           =   8892
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   6960
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
      Picture         =   "lian_add_table.frx":47A6
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   3960
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
      Enabled         =   0   'False
      Picture         =   "lian_add_table.frx":4BF8
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   840
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
      Picture         =   "lian_add_table.frx":504A
   End
   Begin Threed.SSCommand cmdDel 
      Height          =   732
      Left            =   5520
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
      MouseIcon       =   "lian_add_table.frx":549C
      Picture         =   "lian_add_table.frx":58EE
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   2400
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
      MouseIcon       =   "lian_add_table.frx":5D40
      Picture         =   "lian_add_table.frx":6192
   End
   Begin VB.Data Lian_Add_Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\LJXT\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lian_add_table"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1932
   End
End
Attribute VB_Name = "lian_add_table_input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim change_flag As Boolean
Private Sub cmdDel_Click()
    Dim i%
    Dim Temp_record As Integer
    If lian_liao_Data.Recordset.RecordCount <= 0 Then
                    Exit Sub
    End If
    If lian_liao_Data.Recordset.EOF() Then
            Speak_Error ("须选择配方")
        Exit Sub
   End If

If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
            Temp_record = lian_liao_Data.Recordset.Fields("duan").Value
              lian_liao_Data.Recordset.Delete
              i% = 1
              lian_liao_Data.Refresh
              Do While Not lian_liao_Data.Recordset.EOF()
                  lian_liao_Data.Recordset.Edit
                  lian_liao_Data.Recordset.Fields("duan").Value = i%
                  lian_liao_Data.Recordset.Update
                  i% = i% + 1
                  lian_liao_Data.Recordset.MoveNext
            Loop
           lian_liao_Data.Refresh
            With lian_liao_Data.Recordset
                .FindFirst "duan=" & Temp_record
                If .NoMatch Then
                    If .RecordCount > 0 Then
                     .MoveLast
                    End If
                End If
            End With
           
            lian_DBGrid.AllowAddNew = True
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DPei_fan_table.Recordset.Delete
End If

    
End Sub

Private Sub Exit_Command_Click()
    Unload Me
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Dim i As Integer
       Width = MAIN_MDI.Width  ' Set width of form.
       Height = MAIN_MDI.Height + 400
       Left = 0
      Top = 0
      Call Add_Win(Hwnd)
      Lian_Add_Data.DatabaseName = Data_Path & "\ljxt.mdb"

      Lian_Add_Data.RecordSource = "select * from lian_add_table where pei_number like '" & Pei_fan_table_Input.Pei_Fan_Number & "' and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine
      Lian_Add_Data.Refresh

      
      
      If Lian_Add_Data.Recordset.RecordCount <= 0 Then
        Lian_Add_Data.Recordset.AddNew
        Lian_Add_Data.Recordset.Fields("pei_number").Value = Pei_fan_table_Input.Pei_Fan_Number
        Lian_Add_Data.Recordset.Update
        Run_Grid.Refresh
        Run_Grid.AllowAddNew = False
        
      End If
        'For i = 0 To run_Grid.Columns.count - 1
         '      run_Grid.Columns(i).Width = 1200
        'Next i


      lian_liao_Data.DatabaseName = Data_Path & "\ljxt.mdb"
      lian_liao_Data.RecordSource = "select * from lain_liao_table where pei_number like '" & Pei_fan_table_Input.Pei_Fan_Number & "' and mathine= " & Pei_fan_table_Input.Pei_Fan_Mathine & "  order by duan "
      lian_liao_Data.Refresh
      If lian_liao_Data.Recordset.RecordCount <= 0 Then
        lian_liao_Data.Recordset.AddNew
        lian_liao_Data.Recordset.Fields("pei_number").Value = Pei_fan_table_Input.Pei_Fan_Number
        lian_liao_Data.Recordset.Fields("duan").Value = 1
        lian_liao_Data.Recordset.Update
        lian_DBGrid.Refresh
        lian_DBGrid.AllowAddNew = False
      End If
      
            lian_DBGrid.Columns("mathine").Visible = False
            lian_DBGrid.Columns("dou_you_name").Visible = False
            lian_DBGrid.Columns("dou_tan_name").Visible = False
            lian_DBGrid.Columns("dou_jiao_name").Visible = False
            lian_DBGrid.Columns("dou_xiao_name").Visible = False
            lian_DBGrid.Columns("upper_pre").Visible = False
            lian_DBGrid.Columns("upper_statue").Visible = False
            lian_DBGrid.Columns(0).Visible = False
            
            lian_DBGrid.Columns("stop_time").Visible = False
            lian_DBGrid.Columns("duan").Locked = True
            
            For i = 0 To lian_DBGrid.Columns.count - 1
                lian_DBGrid.Columns(i).Width = 600
            Next i
            lian_DBGrid.Refresh
        
    
End Sub





Private Sub Form_Paint()
If China_English = CHINA Then


       Run_Grid.Caption = Pei_fan_table_Input.Pei_Fan_Number & "配方的混练数据"
       Run_Grid.Columns("dou_liao_start").Caption = " 排料时上顶栓状态"
       Run_Grid.Columns("pai_liao_time").Caption = "排料时间"
       Run_Grid.Columns("pai_liao_ask").Caption = " 排料条件"
       Run_Grid.Columns("max_wen_du").Caption = "上限胶温"
       Run_Grid.Columns("min_wen_du").Caption = "下限胶温"
       
       Run_Grid.Columns("temp1").Caption = "返料时间"
       Run_Grid.Columns("temp2").Caption = "小料公差"
       Run_Grid.Columns("temp3").Caption = "备用2"
       
       Run_Grid.Columns("pei_number").Caption = " 编号"
       Run_Grid.Columns("weigh").Caption = " 小料重量"
       Run_Grid.Columns("xiao_liao_name").Caption = "小料名"
    Else
       

       Run_Grid.Caption = " "
       Run_Grid.Columns("dou_liao_start").Caption = " dou"
       Run_Grid.Columns("pai_liao_time").Caption = "pai"
       Run_Grid.Columns("pai_liao_ask").Caption = " T"
       Run_Grid.Columns("max_wen_du").Caption = "X"
       Run_Grid.Columns("min_wen_du").Caption = "S"
       
       Run_Grid.Columns("temp1").Caption = "F"
       Run_Grid.Columns("temp2").Caption = "G"
       Run_Grid.Columns("temp3").Caption = "B2"
       
       Run_Grid.Columns("pei_number").Caption = " N"
       Run_Grid.Columns("weigh").Caption = " X"
       Run_Grid.Columns("xiao_liao_name").Caption = "Q"

    
       
    End If
        
   If China_English = CHINA Then
           
       CmdFind.Caption = "F6 查找"
        cmdDel.Caption = "F10 删除： "
        Insert_Command.Caption = "F3 插入"
        Exit_Command.Caption = "ESC 退出"
        Print_Command.Caption = "F4 打印"
       
        lian_DBGrid.Columns(0).Caption = "B"
        lian_DBGrid.Columns(1).Caption = "段号"

        lian_DBGrid.Columns("dou_you2_time").Caption = "投油二"
        lian_DBGrid.Columns("dou_you_time").Caption = "投油一"
        lian_DBGrid.Columns("dou_tan_time").Caption = "投炭"
        lian_DBGrid.Columns("dou_jiao_time").Caption = "投胶"
        lian_DBGrid.Columns("dou_xiao_time").Caption = "投小料"
        lian_DBGrid.Columns("st").Caption = "投料上升"   '原为上升
        lian_DBGrid.Columns("speed").Caption = "转速"
        lian_DBGrid.Columns("pre").Caption = "上顶栓压力"
        lian_DBGrid.Columns("lian_time").Caption = "混炼时间"
       'lian_DBGrid.Columns(15).Caption = "上顶栓压力："
       'lian_DBGrid.Columns(16).Caption = "上顶栓状态 ："
       lian_DBGrid.Columns("tem").Caption = "温度"
       lian_DBGrid.Columns("neg").Caption = "能量"
       lian_DBGrid.Columns("ctr").Caption = "控制关系 "

       
    Else
       CmdFind.Caption = "F6 search："
        cmdDel.Caption = "F10del： "
        Insert_Command.Caption = "F3inse"
        Exit_Command.Caption = "ESC "
        Print_Command.Caption = "F4 print"
       
       lian_DBGrid.Columns(0).Caption = "B"
       lian_DBGrid.Columns(1).Caption = "D"

        lian_DBGrid.Columns("dou_you2_time").Caption = "time2"
        lian_DBGrid.Columns("dou_you_time").Caption = "time1"
        lian_DBGrid.Columns("dou_tan_time").Caption = "tan"
        lian_DBGrid.Columns("dou_jiao_time").Caption = "jiao"
        lian_DBGrid.Columns("dou_xiao_time").Caption = "xiao"
        lian_DBGrid.Columns("st").Caption = "real"   '原为上升
        lian_DBGrid.Columns("speed").Caption = "speed"
        lian_DBGrid.Columns("pre").Caption = "power"
        lian_DBGrid.Columns("lian_time").Caption = "huan"
       'lian_DBGrid.Columns(15).Caption = "上顶栓压力："
       'lian_DBGrid.Columns(16).Caption = "上顶栓状态 ："
       lian_DBGrid.Columns("tem").Caption = "wen"
       lian_DBGrid.Columns("neg").Caption = "neng"
       lian_DBGrid.Columns("ctr").Caption = "k "

   End If
        
   

End Sub

Private Sub Form_Terminate()
Call Del_Win(Hwnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Del_Win(Hwnd)

End Sub

Private Sub Insert_Command_Click()
    Dim i%
    Dim Temp_record As Integer
    If lian_DBGrid.EditActive = True Then
            Speak_Error ("编辑时不能插入")
            lian_DBGrid.SetFocus
            Exit Sub
    End If

With lian_liao_Data.Recordset
  If .RecordCount > 0 Then
        If .AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
            lian_DBGrid.SetFocus
            Exit Sub
      End If
      temp% = lian_DBGrid.Columns("duan").Value
'        run_Grid.AllowAddNew = True
        i% = temp%
        .FindFirst "duan=" & i%
        Do While Not .NoMatch
                .Edit
                .Fields("duan").Value = i% + 1
                .Update
                i% = i% + 1
                .FindNext "duan=" & i%
        Loop
        Dim strsqlrestore   As String
         strsqlrestore = "insert  into lain_liao_table (mathine,duan,pei_number)  " _
                                          & "  values (" & Pei_fan_table_Input.Pei_Fan_Mathine & " ," & temp% & ",'" & Pei_fan_table_Input.Pei_Fan_Number & "'" & ")"
    
        lian_liao_Data.Database.Execute strsqlrestore, dbFailOnError
        lian_liao_Data.Refresh
        lian_liao_Data.Recordset.FindFirst "duan=" & temp%
        lian_DBGrid.AllowAddNew = False
        'lian_DBGrid.Refresh
  End If
    End With
        lian_DBGrid.SetFocus
        lian_DBGrid.col = 7
        Exit Sub
ErrorHandler:
    Speak_Error ("须选择配方")
    Run_Grid.SetFocus
End Sub

Private Sub lian_DBGrid_AfterUpdate()
        change_flag = UNCHANGE
End Sub

Private Sub lian_DBGrid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
            change_flag = CHANGE
            
End Sub

Private Sub lian_DBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    'If ColIndex = 0 Then
     '   If Right(Trim(lian_DBGrid.Columns(0).Text), 1) < "0" Or Right(Trim(lian_DBGrid.Columns(0).Text), 1) > "9" Then
      '          lian_DBGrid.Columns(0).Text = lian_liao_Data.Recordset.AbsolutePosition + 1
       ' End If
    'End If
    If ColIndex < 1 Then Exit Sub
    With lian_DBGrid
                If Not Check_Float(.Columns(ColIndex).Value) Then
                        .Columns(ColIndex).Value = 0
                End If
    End With
      
End Sub





Private Sub lian_DBGrid_BeforeUpdate(Cancel As Integer)
        lian_DBGrid.Columns("mathine").Text = Pei_fan_table_Input.Pei_Fan_Mathine
End Sub

Private Sub lian_DBGrid_GotFocus()
        lian_DBGrid.col = 7
End Sub

Private Sub lian_DBGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i, Flag As Integer
    Select Case KeyCode
        Case vbKeyUp
              Flag = 1
             For i = 2 To lian_DBGrid.Columns.count - 1
                 If lian_DBGrid.Columns(i).Text <> "" Then
                     Flag = 0
                     Exit For
                 End If
              Next i
               If change_flag = CHANGE Then
                        If Flag = 1 Then KeyCode = 0
                End If
                lian_DBGrid.AllowAddNew = False
        Case 40
            Flag = 1
              For i = 2 To lian_DBGrid.Columns.count - 1
                If lian_DBGrid.Columns(i).Text <> "" And lian_DBGrid.Columns(i).Visible Then
                    Flag = 0
                    Exit For
                End If
              Next i
            
              If lian_DBGrid.Columns(0).Text = "" And lian_DBGrid.Columns(0).Visible = True Then
                    Flag = 1
              End If
              If lian_DBGrid.Columns(1).Text = "" And lian_DBGrid.Columns(1).Visible = True Then
                    Flag = 1
              End If
            
            If Flag = 0 Then
                lian_DBGrid.AllowAddNew = True
            Else
                lian_DBGrid.AllowAddNew = False
            End If
            lian_DBGrid.SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            KeyCode = 0
            lian_DBGrid.SetFocus
        Case vbKeyF3
            Call Insert_Command_Click
            KeyCode = 0
            lian_DBGrid.SetFocus

            
        Case vbKeyF10
            Call cmdDel_Click
            KeyCode = 0
            lian_DBGrid.SetFocus
    End Select
End Sub

Private Sub lian_DBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
              Dim old_index As Integer
              If lian_DBGrid.VisibleRows < 0 Then
                    lian_DBGrid.AllowAddNew = True
                    Exit Sub
              End If
              
             If lian_DBGrid.row <= lian_DBGrid.VisibleRows - 2 Then
                        Exit Sub
             End If
            If lian_DBGrid.col < 2 Then
                    lian_DBGrid.col = 7
            End If
            If lian_DBGrid.Columns("pei_number").Text = "" And lian_DBGrid.Columns("duan").Text = "" Then
                   old_index = lian_liao_Data.Recordset.RecordCount + 1
                   lian_DBGrid.Columns("pei_number").Value = Pei_fan_table_Input.Pei_Fan_Number
                   lian_DBGrid.Columns("duan").Value = old_index
                   'lian_DBGrid.SetFocus
            End If
            If lian_DBGrid.Columns("duan").Value = "" Then
                    old_index = lian_liao_Data.Recordset.AbsolutePosition + 1
                    lian_DBGrid.Columns("duan").Value = old_index
            End If
            
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
  With lian_liao_Data
        Printer.Height = (.Recordset.RecordCount + 6) * Run_Grid.RowHeight
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
    'Printer.FontSize = 20
 
        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
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
          k% = 0
           For i% = 0 To Run_Grid.Columns.count - 1
              If Run_Grid.Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(Run_Grid.Columns(i%).Value); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
              End If
          Next i%
            Printer.Print String(1, " ")
            Y = Printer.CurrentY + 10
            Printer.Line (0, Y)-(Printer.ScaleWidth, Y)     '  print --------
            Printer.Print
        
'打印段标题
   temp(0) = 6
   k = 1
    With lian_DBGrid
   For i = 0 To .Columns.count - 1
         'temp(i) = i * Printer.ScaleWidth / run_Grid.VisibleCols
         If .Columns(i).Visible Then
                temp(k) = 0
                temp(k) = temp(k - 1) + .Columns(i).Width * (Printer.ScaleWidth) / .Width
                k = k + 1
        End If
   Next i
   
    
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    For Font_Control = 1 To .HeadLines
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
       
    '打印段内容
    i% = 0
    With lian_liao_Data
    .Refresh
    
    Do While Not .Recordset.EOF()
        If i% < 48 Then
            k% = 0
            For j% = 0 To lian_DBGrid.Columns.count - 1
               If lian_DBGrid.Columns(j%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print lian_DBGrid.Columns(j%).Value + String(6, " ");
                    k% = k% + 1
               End If
            Next j%
            Printer.Print String(1, " ")
        Else
            Printer.Print "页", Printer.Page
            Printer.NewPage
            i% = 0
        End If
        .Recordset.MoveNext
    Loop
  End With
    'Do While i% < 48
     ''       Printer.Print
       '     i% = i% + 1
    'Loop
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
'    Printer.Print space(10) + "总段数:" & CStr(lian_liao_Data.Recordset.RecordCount)
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





Private Sub Run_Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
            Select Case ColIndex
                Case 1 To 4, 6 To 11
                    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And (KeyAscii <> 46 And KeyAscii <> 8) Then
                        Speak_Error ("必须输入班别1~9")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
                Case 5
                    If KeyAscii > Asc("9") Or KeyAscii < Asc("0") And (KeyAscii <> 46 And KeyAscii <> 8) And KeyAscii <> Asc(".") Then
                        Speak_Error ("必须输入车1~9.")
                        Cancel = True
                        Run_Grid.SetFocus
                    End If
            End Select
            
End Sub

Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    
            
'        Dim temp_max, temp_min   As Integer
'        temp_max = CInt(max_wen_du_text.Text)
'        If min_wen_du_text.Text <> "" Then
 '               temp_min = CInt(min_wen_du_text.Text)
  '      Else
   '             temp_min = 0
    '    End If
     '   If temp_max < temp_min Then
      '      Call Speak_Error("最大温度必须大于最小温度")
       ' End If
    
End Sub


Private Sub run_Grid_BeforeUpdate(Cancel As Integer)
            Run_Grid.Columns("mathine").Text = Pei_fan_table_Input.Pei_Fan_Mathine
End Sub

Private Sub run_Grid_GotFocus()
            Run_Grid.col = 1
End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, Flag As Integer
    Select Case KeyCode
        Case vbKeyF10
            KeyCode = 0
            Run_Grid.SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            KeyCode = 0
            Run_Grid.SetFocus
            
    End Select
    
End Sub




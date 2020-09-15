VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Pei_fan_table_Input 
   Caption         =   "Pei_fan_Table_Input"
   ClientHeight    =   5820
   ClientLeft      =   300
   ClientTop       =   828
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   8880
   Begin Threed.SSFrame SSFrame1 
      Height          =   1092
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   8652
      _Version        =   65536
      _ExtentX        =   15261
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
      ShadowStyle     =   1
      Begin Threed.SSCommand Tan_Hei_Command 
         Height          =   612
         Left            =   360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   1080
         _StockProps     =   78
         Caption         =   "Tan"
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
      End
      Begin Threed.SSCommand You_Liao_Command 
         Height          =   612
         Left            =   2040
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   1080
         _StockProps     =   78
         Caption         =   "You"
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
      End
      Begin Threed.SSCommand Huan_Lian_Command 
         Height          =   612
         Left            =   3720
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   1080
         _StockProps     =   78
         Caption         =   "Hun"
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
      End
      Begin Threed.SSCommand Jiao_Liao_Command 
         Height          =   612
         Left            =   5400
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   1080
         _StockProps     =   78
         Caption         =   "Jiao"
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
      End
      Begin Threed.SSCommand You2_Command 
         Height          =   612
         Left            =   7080
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   1080
         _StockProps     =   78
         Caption         =   "You"
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
      End
   End
   Begin VB.Data DPei_fan_table 
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
      RecordSource    =   "pei_fan_table"
      Top             =   5280
      Visible         =   0   'False
      Width           =   912
   End
   Begin MSDBGrid.DBGrid run_Grid 
      Bindings        =   "Pei_fan_table_Input.frx":0000
      Height          =   3048
      Left            =   120
      OleObjectBlob   =   "Pei_fan_table_Input.frx":0019
      TabIndex        =   0
      Top             =   960
      Width           =   8652
   End
   Begin Threed.SSCommand Exit_Command 
      Height          =   732
      Left            =   6960
      TabIndex        =   1
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
      Font3D          =   4
      Picture         =   "Pei_fan_table_Input.frx":189C
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   3840
      TabIndex        =   2
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
      Font3D          =   4
      Picture         =   "Pei_fan_table_Input.frx":1CEE
   End
   Begin Threed.SSCommand Insert_Command 
      Height          =   732
      Left            =   720
      TabIndex        =   3
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
      Font3D          =   4
      Picture         =   "Pei_fan_table_Input.frx":2140
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   732
      Left            =   5400
      TabIndex        =   4
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
      Font3D          =   4
      MouseIcon       =   "Pei_fan_table_Input.frx":2592
      Picture         =   "Pei_fan_table_Input.frx":29E4
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   2280
      TabIndex        =   5
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
      Font3D          =   4
      MouseIcon       =   "Pei_fan_table_Input.frx":2E36
      Picture         =   "Pei_fan_table_Input.frx":3288
   End
End
Attribute VB_Name = "Pei_fan_table_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msBookMark As String

Dim gsDBName As Database
Dim gsThisSet As Recordset
Public You_Liao_Flag  As Integer
Public Pei_Fan_Number  As String
Public Pei_Fan_Mathine As Integer
Dim Pei_Having(0 To 80, 1 To 2) As Variant '配方数据  配方号，机号

Public Sub Add_list()
    Dim number_str    As String * 8, name_str As String * 15
    Dim Mathine%, i%
    Dim temp_index   As Integer

    DPei_fan_table.Refresh
    
    For i = 0 To UBound(Pei_Having)
            Pei_Having(i, 2) = ""
            Pei_Having(i, 1) = ""
    Next i
    
    If Not DPei_fan_table.Recordset.EOF() Then
         Do While Not DPei_fan_table.Recordset.EOF
                        If Check_Float(DPei_fan_table.Recordset.Fields("mathine").Value) Then
                                Mathine = DPei_fan_table.Recordset.Fields("mathine").Value
                                temp_index = Val(DPei_fan_table.Recordset.Fields("index").Value)
                                Pei_Having(temp_index, Mathine) = vFieldVal(DPei_fan_table.Recordset.Fields("_number").Value)
                        End If
                        DPei_fan_table.Recordset.MoveNext
        Loop
        DPei_fan_table.Recordset.MoveFirst
    End If
End Sub









Private Sub cmdDelete_Click()
Dim tBook As String
Dim i%
Dim strsqlrestore  As String
    If DPei_fan_table.Recordset.RecordCount <= 0 Then
                    Exit Sub
    End If
            If DPei_fan_table.Recordset.AbsolutePosition < 0 Then
                    Speak_Error ("须选择配方")
                    Exit Sub
            End If
If China_English = CHINA Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
        cmdDelete.Enabled = False
        
         i% = run_Grid.Columns(0).Value
        strsqlrestore = "delete  from   tan_hei_table    where [pei_number] ='" & run_Grid.Columns("_number").Value & "'"
        DPei_fan_table.Database.Execute strsqlrestore, dbFailOnError
        
        strsqlrestore = "delete  from   you_liao_table    where [pei_number] ='" & run_Grid.Columns("_number").Value & "'"
        DPei_fan_table.Database.Execute strsqlrestore, dbFailOnError
        
        strsqlrestore = "delete  from   lian_add_table    where [pei_number] ='" & run_Grid.Columns("_number").Value & "'"
        Debug.Print strsqlrestore
        DPei_fan_table.Database.Execute strsqlrestore, dbFailOnError
        
        strsqlrestore = "delete  from   jiao_liao_table    where [pei_number] ='" & run_Grid.Columns("_number").Value & "'"
        DPei_fan_table.Database.Execute strsqlrestore, dbFailOnError
        
        strsqlrestore = "delete  from   lain_liao_table    where [pei_number] ='" & run_Grid.Columns("_number").Value & "'"
        DPei_fan_table.Database.Execute strsqlrestore, dbFailOnError
        
        DPei_fan_table.Recordset.Delete
        run_Grid.AllowAddNew = False
        
        
        strsqlrestore = "update pei_fan_table  set  index=index-1   where index>=" & i%
        DPei_fan_table.Database.Execute strsqlrestore, dbFailOnError
        Call Add_list
        DPei_fan_table.Recordset.FindFirst "index=" & i%
        If DPei_fan_table.Recordset.NoMatch Then
            If DPei_fan_table.Recordset.RecordCount > 0 Then
                 DPei_fan_table.Recordset.MoveLast
            End If
        End If
        cmdDelete.Enabled = True
        'run_Grid.Refresh
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DPei_fan_table.Recordset.Delete
End If
    'Call Add_list
End Sub


Private Sub CmdFind_Click()
On Error GoTo FindErr
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
  
FindStart:
  For i = 1 To 4
      frmFind.lstFields.AddItem run_Grid.Columns(i).Caption
  Next


  gbFindFailed = False
  gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOp
  frmFind.txtExpression.Text = gsFindExpr
  frmFind.optFindType(gnFindType) = True

  frmFind.Show vbModal

  
  
 If gbFindFailed = True Then Exit Sub 'find cancelled
 
  sTmp = "[" + run_Grid.Columns(gsFindFiel).DataField + "]" + Space(1) + gsFindOp + Space(1) + "'" & Trim(gsFindExpr) & "'"
  
   sBookMark = DPei_fan_table.Recordset.Bookmark

  Select Case gnFindType
    Case 0
      DPei_fan_table.Recordset.FindFirst sTmp
    Case 1
      DPei_fan_table.Recordset.FindNext sTmp
    Case 2
      DPei_fan_table.Recordset.FindPrevious sTmp
    Case 3
      DPei_fan_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = DPei_fan_table.Recordset.NoMatch
  If mbNotFound = True Then MsgBox "not found!"
AfterWhile:

  Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    DPei_fan_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    DPei_fan_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = DPei_fan_table.Recordset.Bookmark  'save the new position
  
    DPei_fan_table.Recordset.Bookmark = sBookMark
  End If

  mbDataChanged = False

  Exit Sub
FindErr:
    mbNotFound = True
    Exit Sub
End Sub




















Private Sub DPei_fan_table_Validate(Action As Integer, Save As Integer)
      If Save = -1 Then
          For i% = 0 To 7
            If run_Grid.Columns(i%).Text = "" And run_Grid.Columns(i%).Visible Then
                Call Speak_Error("应输入" & run_Grid.Columns(i%).Caption)
                Save = 0
                  run_Grid.SetFocus
                Exit Sub
         End If
        Next i%
    End If
End Sub


Private Sub Exit_Command_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
        Me.WindowState = 2
        run_Grid.SetFocus
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
            Width = MAIN_MDI.Width  ' Set width of form.
            Height = MAIN_MDI.Height - 600
            Left = 0
            Top = 0

             Call Add_Win(Hwnd)
            
           DPei_fan_table.DatabaseName = Data_Path & "\ljxt.mdb"
           DPei_fan_table.RecordSource = "select * from   pei_fan_table order by  index"
           DPei_fan_table.Refresh
           
      If DPei_fan_table.Recordset.RecordCount <= 0 Then
                DPei_fan_table.Recordset.AddNew
                 DPei_fan_table.Recordset.Fields("index").Value = 1
                 DPei_fan_table.Recordset.Fields("m_day").Value = Format(Date$, "yy.mm.dd")
                 DPei_fan_table.Recordset.Fields("m_time").Value = Format(time$, "hh:mm")
                 DPei_fan_table.Recordset.Fields("flag").Value = "0"
                 
                 DPei_fan_table.Recordset.Fields("mathine").Value = IIf(Make_Mathine = 2, 2, 1)
                 DPei_fan_table.Recordset.Update
                 DPei_fan_table.Refresh
        End If
               Call Add_list
                DPei_fan_table.Refresh
               run_Grid.Refresh
               
End Sub






Private Sub Form_Paint()
      If China_English = CHINA Then
          cmdDelete.Caption = "F10删除配方"
          Insert_Command.Caption = "F3插入"
          CmdFind.Caption = "F6查询配方"
          Exit_Command.Caption = "ESC退出"
          Print_Command.Caption = "F4打印"
          Me.Caption = "配方目录输入:"

            run_Grid.Caption = "配方表"
          run_Grid.Columns(0).Caption = "目录序号:"
          run_Grid.Columns(1).Caption = "配方编号:"
          run_Grid.Columns(2).Caption = "配方名称:"
          run_Grid.Columns(3).Caption = "日期:"
          run_Grid.Columns(4).Caption = "时间:"
          run_Grid.Columns("mathine").Caption = "机号:"

          Tan_Hei_Command.Caption = "F7碳黑录入"
          You_Liao_Command.Caption = "F8油料录入"
          Huan_Lian_Command.Caption = "F5混炼录入"
          Jiao_Liao_Command.Caption = "F9胶料录入"
          You2_Command.Caption = "F2油料2录入"


        Else

          run_Grid.Caption = "PEI-FAN_TAB"
          cmdDelete.Caption = "F10Del"
          CmdFind.Caption = "&Search"
          Insert_Command.Caption = "F3Insert"
          Exit_Command.Caption = "&Exit"
          Print_Command.Caption = "&Print"
          Me.Caption = "pei_fan_table_input:"

          
          run_Grid.Columns(0).Caption = "index:"
          run_Grid.Columns(1).Caption = "code:"
          run_Grid.Columns(2).Caption = "name:"
          run_Grid.Columns(3).Caption = "date:"
          run_Grid.Columns(4).Caption = "time:"
          run_Grid.Columns("mathine").Caption = "mathine:"

          Tan_Hei_Command.Caption = "F7 tan hei"
          You_Liao_Command.Caption = "F8 you liao"
          Huan_Lian_Command.Caption = "F5 huan lian"
          Jiao_Liao_Command.Caption = "F9 jiao liao "
          You2_Command.Caption = "F2 you liao 2"


        End If
End Sub

Private Sub Form_Terminate()
Call Del_Win(Hwnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
        Dim strsqlrestore   As String
        strsqlrestore = "delete  from   pei_fan_table    where [_number] is null"
        DPei_fan_table.Database.Execute strsqlrestore, dbFailOnError
        Call Del_Win(Hwnd)

End Sub

Private Sub huan_lian_command_Click()
    Dim Temp_Str As String
    For i% = 0 To 10
            If run_Grid.Columns(i%).Text = "" And run_Grid.Columns(i%).Visible = True Then
                Call Speak_Error("应输入" & run_Grid.Columns(i%).Caption)
                If i% > 0 Then run_Grid.SetFocus
                Exit Sub
             End If
    Next i%
    Temp_Str = run_Grid.Columns("_number").Text
    
    
    If Password_Do("lian_add_table_Input", Mid(Huan_Lian_Command.Caption, 3)) = False Then
        run_Grid.SetFocus
        Exit Sub
    End If
    If Temp_Str = "" Then
            MsgBox "必须有配方编号"
    Else
        Pei_Fan_Number = Temp_Str
        Pei_Fan_Mathine = run_Grid.Columns("mathine").Text
        lian_add_table_input.Show
         If China_English = CHINA Then
            lian_add_table_input.Caption = "密炼工艺    配方名：" & run_Grid.Columns("name").Text
          Else
      End If
        
   End If

End Sub



Private Sub Insert_Command_Click()
 Dim temp%
 Dim strsqlrestore   As String
 On Error GoTo ErrorHandler:
 
 If run_Grid.EditActive = True Then
            Speak_Error ("编辑时不能插入")
            run_Grid.SetFocus
            Exit Sub
 End If
 
  If DPei_fan_table.Recordset.RecordCount > 0 Then
        If DPei_fan_table.Recordset.AbsolutePosition < 0 Then
            Speak_Error ("须选择插入位置")
        Exit Sub
      End If
      temp% = run_Grid.Columns(0).Value
'        run_Grid.AllowAddNew = True
        
        strsqlrestore = "update pei_fan_table  set  index=index+1   where index>=" & temp%
        Debug.Print strsqlrestore
        DPei_fan_table.Database.Execute strsqlrestore, dbFailOnError
        
'        i% = temp%
 '       DPei_fan_table.Recordset.FindFirst "index=" & i%
  '      Do While Not DPei_fan_table.Recordset.NoMatch
   '             DPei_fan_table.Recordset.Edit
    '            DPei_fan_table.Recordset.Fields("index").value = i% + 1
     '           DPei_fan_table.Recordset.Update
      '          i% = i% + 1
       '         DPei_fan_table.Recordset.FindNext "index=" & i%
        'Loop
        
                DPei_fan_table.Recordset.AddNew
                DPei_fan_table.Recordset.Fields("index").Value = temp%
                 DPei_fan_table.Recordset.Fields("m_day").Value = Format(Date$, "yy.mm.dd")
                 DPei_fan_table.Recordset.Fields("m_time").Value = Format(time$, "hh:mm")
                 DPei_fan_table.Recordset.Fields("flag").Value = 0
                 DPei_fan_table.Recordset.Fields("mathine").Value = Make_Mathine
        
        DPei_fan_table.Recordset.Update
        DPei_fan_table.Refresh
        DPei_fan_table.Recordset.FindFirst "index=" & temp%
        run_Grid.Refresh
        run_Grid.AllowAddNew = False
  End If
        run_Grid.SetFocus
        
  Exit Sub
ErrorHandler:
    Speak_Error ("错误：Insert ")
    run_Grid.SetFocus

End Sub

Private Sub jiao_liao_command_Click()
    Dim Temp_Str As String
        For i% = 0 To 10
            If run_Grid.Columns(i%).Text = "" And run_Grid.Columns(i%).Visible = True Then
                Call Speak_Error("应输入" & run_Grid.Columns(i%).Caption)
                If i% > 0 Then run_Grid.SetFocus
                Exit Sub
             End If
    Next i%
    Temp_Str = run_Grid.Columns("_number").Text
    
    If Password_Do("Jiao_liao_table", Mid(Jiao_Liao_Command.Caption, 3)) = False Then
        run_Grid.SetFocus
        Exit Sub
    End If
    If Temp_Str = "" Then
            MsgBox "必须有配方编号"
    Else
        Pei_Fan_Number = Temp_Str
        Pei_Fan_Mathine = run_Grid.Columns("mathine").Text
        jiao_liao_table_Input.Show
        jiao_liao_table_Input.Caption = "胶料工艺    配方名：" & run_Grid.Columns("name").Text
   End If
End Sub









Private Sub Print_Command_Click()
'   On Error GoTo error_doing
   Dim Y  As Long
   Dim i%, k%
   Dim temp(0 To 20)      As Long
   Dim book   As Variant
   book = run_Grid.Bookmark
   
   Screen.MousePointer = vbHourglass

   If Printer.ScaleWidth < run_Grid.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        run_Grid.SetFocus
        Exit Sub
    End If
  With DPei_fan_table
        Printer.Height = (.Recordset.RecordCount + 6) * run_Grid.RowHeight
        Printer.Width = Screen.Width
   End With
    
      temp(0) = 6
   k = 1
   With run_Grid
   For i = 0 To .Columns.count - 1
         'temp(i) = i * Printer.ScaleWidth / run_Grid.VisibleCols
         If .Columns(i).Visible Then
                temp(k) = 0
                temp(k) = temp(k - 1) + .Columns(i).Width * (Printer.ScaleWidth) / .Width
                k = k + 1
        End If
   Next i
   End With
    
   
    Printer.Print
    Printer.Print
    'Printer.FontName = "Courier"
    'Printer.FontSize = 18
    Dim CX
    CX = Printer.CurrentY
        Printer.PSet (Printer.ScaleWidth / 3, CX), 1
        Printer.Print run_Grid.Caption + Space(4) + Date$
        
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    'Printer.FontName = "Courier"
    'Printer.FontSize = 12
    
        k% = 0
        For i% = 0 To run_Grid.Columns.count - 1
            If run_Grid.Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(run_Grid.Columns(i%).Caption); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
            End If
        Next i%
    
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    DPei_fan_table.Refresh
    j% = 0
    Do While Not DPei_fan_table.Recordset.EOF()
        If j% < 48 Then
            k% = 0
           For i% = 0 To run_Grid.Columns.count - 1
              If run_Grid.Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(run_Grid.Columns(i%).Value); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
              End If
          Next i%
         Printer.Print String(1, " ")
         DPei_fan_table.Recordset.MoveNext
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
    'Printer.Print Space(10) + "总数:" & CStr(DPei_fan_table.Recordset.RecordCount)
    'Printer.Print "第" + CStr(Printer.Page) + "页"
    Printer.EndDoc
    Screen.MousePointer = vbDefault
'    Run_Grid.Bookmark = book
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
      run_Grid.Bookmark = book

End Sub















Private Sub run_Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
            Dim i%, temp_local%
            Dim Mathine%
            Select Case ColIndex
                Case 1
                    If run_Grid.Columns(1).Text = "" Then
                        Speak_Error ("配方编号长度>=1")
                        Cancel = True
                        run_Grid.SetFocus
                        Exit Sub
                    End If
                    If Len(run_Grid.Columns(1).Text) > 6 Then
                        Speak_Error ("配方编号长度必须《6")
                        Cancel = True
                        run_Grid.SetFocus
                        Exit Sub
                    End If
                Mathine = Val(run_Grid.Columns("mathine").Value)
                If Mathine = 0 Then Exit Sub
                temp_local% = Val(run_Grid.Columns("index").Value)
              For i% = 0 To UBound(Pei_Having)
                If Trim(Pei_Having(i%, Mathine)) = run_Grid.Columns("_number").Value And i% <> temp_local% Then
                        Speak_Error (str(Mathine) + "机组的" + run_Grid.Columns("_number").Caption & "重")
                        Cancel = True
                        Exit Sub
                End If
            Next i%
            If temp_local% < 0 Then
                temp_local% = 0
            End If
            If temp_local% > UBound(Pei_Having) Then
                        Speak_Error (str(Mathine) + "机组的" + "配方数必须《" + UBound(pei_havin))
                        Cancel = True
                        Exit Sub
                End If
                Pei_Having(temp_local%, Mathine) = run_Grid.Columns("_number").Value
                Case 2
                    If run_Grid.Columns(2).Text = "" Then
                        Speak_Error ("配方名称长度>=1")
                        Cancel = True
                        run_Grid.SetFocus
                    End If
                    If Len(run_Grid.Columns(2).Text) > 15 Then
                        Speak_Error ("配方名称长度必须《15")
                        Cancel = True
                                run_Grid.SetFocus
                    End If
                Case 10
                    If (run_Grid.Columns(10).Text = "") And run_Grid.Columns(10).Visible = True Then
                        run_Grid.Columns(10).Value = 1
                    End If
                    If run_Grid.Columns(10).Text <> "2" And run_Grid.Columns(10).Text <> "1" Then
                        Speak_Error ("必须输入机号1~2")
                        Cancel = True
                    End If
                    run_Grid.SetFocus
                End Select

End Sub


Private Sub run_Grid_GotFocus()
        If run_Grid.col = 0 Then
                run_Grid.col = 1
        End If

End Sub

Private Sub run_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, Flag As Integer
    Select Case KeyCode
        Case 40
            Flag = 0
            For i = 0 To 10
              If run_Grid.Columns(i).Value = "" And run_Grid.Columns(i).Visible = True Then
                    Flag = 1
                    Exit For
              End If
            Next i
            If Flag = 0 Then
                run_Grid.AllowAddNew = True
            Else
                run_Grid.AllowAddNew = False
            End If
         Case vbKeyEscape
                Unload Me
             
        Case vbKeyF5
            run_Grid.EditActive = False
            Call huan_lian_command_Click
            
        Case vbKeyF7
            run_Grid.EditActive = False
            Call tan_hei_command_Click
        Case vbKeyF2
            run_Grid.EditActive = False
            Call You2_Command_Click
            'run_Grid.SetFocus
        Case vbKeyF8
            run_Grid.EditActive = False
            Call you_liao_command_Click
            
        Case vbKeyF9
            run_Grid.EditActive = False
            Call jiao_liao_command_Click
            
        Case vbKeyF10
            Call cmdDelete_Click
            KeyCode = 0
            run_Grid.SetFocus
        Case vbKeyF3
            Call Insert_Command_Click
            run_Grid.SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            run_Grid.SetFocus
            
        Case vbKeyF6
            Call CmdFind_Click
            run_Grid.SetFocus
            
    End Select
'    Run_Grid.SetFocus
End Sub

Private Sub run_Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
              Dim old_index As Integer
           '   On Error GoTo errorHandle
              
              If run_Grid.VisibleRows <= 1 Then
                    Exit Sub
              End If
        If run_Grid.col = 0 And LastCol > 0 Then
                run_Grid.col = LastCol
        End If
              
            If run_Grid.Columns("index").Value = "" Then
                 old_index = DPei_fan_table.Recordset.RecordCount + 1
                 run_Grid.Columns("index").Value = old_index
                 run_Grid.Columns("m_day").Text = Format(Date$, "yy.mm.dd")
                 run_Grid.Columns("m_time").Text = Format(time$, "hh:mm")
                 run_Grid.Columns("flag").Text = "0"
                 run_Grid.Columns("mathine").Text = IIf(Make_Mathine = 2, 2, 1)
            End If
            Exit Sub
errorHandle:
        
        run_Grid.SetFocus
        
End Sub



Private Sub tan_hei_command_Click()
    Dim Temp_Str As String
    For i% = 0 To 10
            If run_Grid.Columns(i%).Text = "" And run_Grid.Columns(i%).Visible = True Then
                Call Speak_Error("应输入" & run_Grid.Columns(i%).Caption)
                If i% > 0 Then run_Grid.SetFocus
                Exit Sub
     End If
Next i%
    Temp_Str = run_Grid.Columns("_number").Text
    If Password_Do("Tan_hei_table_Input", Mid(Tan_Hei_Command.Caption, 3)) = False Then
        run_Grid.SetFocus
        Exit Sub
    End If
        Pei_Fan_Number = Temp_Str
        Pei_Fan_Mathine = run_Grid.Columns("mathine").Text
        Tan_hei_table_Input.Show
        Tan_hei_table_Input.ZOrder 0
        
         If China_English = CHINA Then
            Tan_hei_table_Input.Caption = "炭黑输入    配方名：" & run_Grid.Columns("name").Text
          Else
      End If
              Tan_hei_table_Input.Show
        'Tan_hei_table_Input.ZOrder 0

End Sub




Private Sub you_liao_command_Click()
    Dim Temp_Str As String
    For i% = 0 To 10
            If run_Grid.Columns(i%).Text = "" And run_Grid.Columns(i%).Visible = True Then
                Call Speak_Error("应输入" & run_Grid.Columns(i%).Caption)
                If i% > 0 Then run_Grid.SetFocus
                Exit Sub
             End If
    Next i%
    Temp_Str = run_Grid.Columns("_number").Text
    
    If Password_Do("You_liao_table_Input", Mid(You_Liao_Command.Caption, 3)) = False Then
        run_Grid.SetFocus
        Exit Sub
    End If
        You_Liao_Flag = 1
        Pei_Fan_Number = Temp_Str
        Pei_Fan_Mathine = run_Grid.Columns("mathine").Text
        You_liao_table_Input.Show
        If China_English = CHINA Then
            You_liao_table_Input.Caption = "油料输入    配方名：" & run_Grid.Columns("name").Text
          Else
      End If
                
End Sub

Private Sub You2_Command_Click()
    Dim Temp_Str As String
    For i% = 0 To 10
            If run_Grid.Columns(i%).Text = "" And run_Grid.Columns(i%).Visible = True Then
                Call Speak_Error("应输入" & run_Grid.Columns(i%).Caption)
                If i% > 0 Then run_Grid.SetFocus
                Exit Sub
             End If
    Next i%
    Temp_Str = run_Grid.Columns("_number").Text
    
    If Password_Do("You2_Table_Input", Mid(You2_Command.Caption, 3)) = False Then
        run_Grid.SetFocus
        Exit Sub
    End If
        You_Liao_Flag = 2
        Pei_Fan_Number = Temp_Str
        Pei_Fan_Mathine = run_Grid.Columns("mathine").Text
        You_liao_table_Input.Show
        If China_English = CHINA Then
            You_liao_table_Input.Caption = "油料秤2输入    配方名：" & run_Grid.Columns("name").Text
          Else
      End If
End Sub

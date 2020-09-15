VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form F_Error_History 
   Caption         =   "错误历史窗口"
   ClientHeight    =   4716
   ClientLeft      =   1176
   ClientTop       =   1200
   ClientWidth     =   7356
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4716
   ScaleWidth      =   7356
   Begin Threed.SSCommand SSCommand1 
      Cancel          =   -1  'True
      Height          =   492
      Left            =   2664
      TabIndex        =   1
      Top             =   4080
      Width           =   1092
      _Version        =   65536
      _ExtentX        =   1926
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "退出(ESC)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   492
      Left            =   4320
      TabIndex        =   2
      Top             =   4080
      Width           =   1092
      _Version        =   65536
      _ExtentX        =   1926
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "帮助(&H)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand Print_Command 
      Default         =   -1  'True
      Height          =   492
      Left            =   6096
      TabIndex        =   3
      Top             =   4080
      Width           =   1092
      _Version        =   65536
      _ExtentX        =   1926
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "打印(&P)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   144
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Error_History_Table"
      Top             =   4056
      Visible         =   0   'False
      Width           =   1932
   End
   Begin MSDBGrid.DBGrid Error_Grid 
      Bindings        =   "F_ErrHis.frx":0000
      Height          =   3852
      Left            =   120
      OleObjectBlob   =   "F_ErrHis.frx":0010
      TabIndex        =   0
      Top             =   120
      Width           =   7092
   End
End
Attribute VB_Name = "F_Error_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：显示过去一段时间内产生的错误信息
Private Sub Error_Grid_DblClick()
Call SSCommand2_Click
End Sub

Private Sub Form_Activate()
Dim i As Integer
Dim Cur_rec As Variant
Cur_rec = Data1.Recordset.Bookmark
i = 0
Data1.Recordset.MoveFirst
While Not Data1.Recordset.EOF
    i = i + 1
    Error_Grid.Columns("index").Value = i
    Data1.Recordset.MoveNext
    Error_Grid.Refresh
Wend
'Data1.Recordset.MoveLast
On Error Resume Next
Data1.Recordset.Bookmark = Cur_rec
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
Error_Grid.Columns(0).Width = 815
Error_Grid.Columns(1).Width = 912
Error_Grid.Columns(2).Width = 912 * 3
Error_Grid.Columns(3).Width = 984
Error_Grid.Columns(4).Width = 972
Error_Grid.Columns(5).Visible = False
Error_Grid.Columns(6).Visible = False
Error_Grid.Columns(7).Visible = False
Error_Grid.Columns(8).Visible = False
Error_Grid.Columns(9).Visible = False
Error_Grid.Columns(10).Visible = False
End Sub

Private Sub Form_Paint()
   If China_English = CHINA Then
        SSCommand1.Caption = "&Esc退出"
        SSCommand2.Caption = "&H帮助"
        Print_Command.Caption = "&P打印"
    Else
        SSCommand1.Caption = "&Esc"
        SSCommand2.Caption = "&Help"
        Print_Command.Caption = "&Print"
        
    End If
End Sub

Private Sub Print_Command_Click()
Dim Y  As Long
   Dim i%, k%
   Dim temp(0 To 20)     As Long
   Dim book   As Variant
   book = Error_Grid.Bookmark
   
   Screen.MousePointer = vbHourglass
    If Printer.ScaleWidth < Error_Grid.Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        Error_Grid.SetFocus
        Exit Sub
    End If
  With Data1
        Printer.Height = (.Recordset.RecordCount + 2) * Error_Grid.RowHeight
        Printer.Width = Screen.Width
   End With
    
      temp(0) = 6
   k = 1
   With Error_Grid
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
        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY), 1
        Printer.Print Error_Grid.Caption + space(4) + Date$
    Else
        Printer.Print String(20, " ");
        k% = 0
        For i% = 0 To Error_Grid.Columns.count - 1
            If Error_Grid.Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(Error_Grid.Columns(i%).Caption); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
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
        For i% = 0 To Error_Grid.Columns.count - 1
            If Error_Grid.Columns(i%).Visible = True Then
                   'Printer.Print SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                   Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(Error_Grid.Columns(i%).Caption); 'SetRightString(Error_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
            End If
        Next i%
    
    Printer.Print String(1, " ")
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    Data1.Recordset.MoveFirst
    j% = 0
    Do While Not Data1.Recordset.EOF()
        If j% < 48 Then
            k% = 0
           For i% = 0 To Error_Grid.Columns.count - 1
              If Error_Grid.Columns(i%).Visible = True Then
'                    Printer.Print SetRightString(Run_Grid.Columns(i%).value, temp(K%));
                    Printer.Print Trim(Error_Grid.Columns(i%).Value); 'SetRightString(Run_Grid.Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    
              End If
          Next i%
         Printer.Print String(1, " ")
         Data1.Recordset.MoveNext
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
    Error_Grid.Bookmark = book
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
      Error_Grid.Bookmark = book
End Sub

Private Sub SSCommand1_Click()
     Unload Me
    'Me.Hide
End Sub

Private Sub SSCommand2_Click()
Dim Error_Code, Error_Prompt As String
With F_ShowLc
    If Not F_Error_Help.Visible Then
        Error_Code = Trim(Error_Grid.Columns(1).Value)
        Error_Prompt = "(" + Error_Code + ") " + Trim(Error_Grid.Columns(2).Value)
        
        .ErrorData.Recordset.index = "error_idx"
        .ErrorData.Recordset.Seek "=", Error_Code
        If Not .ErrorData.Recordset.NoMatch Then
            F_Error_Help.Label2.Caption = Error_Prompt
            If Not IsNull(.ErrorData.Recordset!China_Help) Then
                F_Error_Help.Text1.Text = .ErrorData.Recordset!China_Help
            Else
                F_Error_Help.Text1.Text = ""
            End If
            F_Error_Help.Show vbModal
            Reading_ErrorHelp = True
        Else
            F_Error_Help.Label2.Caption = Error_Prompt
            F_Error_Help.Text1.Text = ""
        End If
    End If
End With
End Sub


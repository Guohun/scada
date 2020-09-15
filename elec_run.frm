VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Elec_Run 
   Caption         =   "Form1"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9276
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5172
   ScaleWidth      =   9276
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   0
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4332
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "elec_run.frx":0000
      TabIndex        =   0
      Tag             =   """0"""
      Top             =   720
      Width           =   4452
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4332
      Index           =   1
      Left            =   4680
      OleObjectBlob   =   "elec_run.frx":09B5
      TabIndex        =   1
      Tag             =   """1"""
      Top             =   720
      Width           =   4452
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   7440
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
      Font3D          =   2
      Picture         =   "elec_run.frx":136A
   End
   Begin Threed.SSCommand CmdFind 
      Height          =   732
      Left            =   5400
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
      Font3D          =   2
      Picture         =   "elec_run.frx":17BC
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   3600
      TabIndex        =   4
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
      MouseIcon       =   "elec_run.frx":1C0E
      Picture         =   "elec_run.frx":2060
   End
   Begin Threed.SSCommand Change_Command 
      Height          =   732
      Left            =   1080
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "NoEdit"
      Top             =   0
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2346
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "Save"
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
      Picture         =   "elec_run.frx":24B2
   End
End
Attribute VB_Name = "Elec_Run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data1(0 To 1) As Database
Dim rec(0 To 1)   As Recordset
Dim ChangeFlag   As Boolean
Dim Root_Change  As Boolean
Private mTotalRows&(0 To 1) ' Contains the total rows in the set of records
Private Light_Data() As Variant ' 2-dimensional array containing records 临时数
Private Check_Data() As Variant ' 2-dimensional array containing records 临时数

'修改电信号命令
Private Sub Change_Command_Click()
    Dim i%, X%
    
    If Change_Command.Tag = "N" Then
            Change_Command.Tag = "Y"
            Change_Command.Caption = "F3调试状态"
            Elec_Output.Flag = 2
             
            Call write_elec_output(Elec_Output)
            If Root_Change Then
                    DBGrid1(0).Columns(7).Locked = False
            End If
            DBGrid1(1).Columns(7).Locked = False
    Else
        If DBGrid1(1).EditActive = True Then
                    Speak_Error (" ")
        End If
            Elec_Output.Flag = 3
            Call write_elec_output(Elec_Output)
           Change_Command.Tag = "N"
           Change_Command.Caption = "F3正常状态"
           DBGrid1(1).Columns(7).Locked = True
            DBGrid1(0).Columns(7).Locked = True
    End If
End Sub


'发现命令
Private Sub CmdFind_Click()
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
  On Error Resume Next
FindStart:
  
   
   frmFind.lstFields.AddItem DBGrid1(0).Columns(0).Caption
   frmFind.lstFields.AddItem DBGrid1(0).Columns(1).Caption
   frmFind.lstFields.AddItem DBGrid1(0).Columns(2).Caption
   frmFind.lstFields.AddItem DBGrid1(0).Columns(3).Caption
    
  'reset the flags
  gbFindFailed = False
  'gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOp
  frmFind.txtExpression.Text = gsFindExpr
  'frmFind.optFindType(gnFindType) = True

  frmFind.Show vbModal

  
  
 If gbFindFailed = True Then Exit Sub 'find ccelled
    For i = 0 To mTotalRows&(0) - 1
            Select Case gsFindFiel
                Case DBGrid1(0).Columns(0).Caption
                            If Light_Data(0, i) = gsFindExpr Then
                                    Exit For
                            End If
                Case DBGrid1(0).Columns(1).Caption
                            If Light_Data(1, i) = gsFindExpr Then
                                    Exit For
                            End If
                Case DBGrid1(0).Columns(2).Caption
                            If Light_Data(2, i) = gsFindExpr Then
                                    Exit For
                            End If
                Case DBGrid1(0).Columns(3).Caption
                            If Light_Data(3, i) = gsFindExpr Then
                                    Exit For
                            End If
                End Select
    Next i
    If i >= mTotalRows&(0) Then
        For i = 0 To mTotalRows&(1) - 1
                Select Case gsFindFiel
                Case DBGrid1(0).Columns(0).Caption
                            If Check_Data(0, i) = gsFindExpr Then
                                    Exit For
                            End If
                Case DBGrid1(0).Columns(1).Caption
                            If Check_Data(1, i) = gsFindExpr Then
                                    Exit For
                            End If
                Case DBGrid1(0).Columns(2).Caption
                            If Check_Data(2, i) = gsFindExpr Then
                                    Exit For
                            End If
                Case DBGrid1(0).Columns(3).Caption
                            If Check_Data(3, i) = gsFindExpr Then
                                    Exit For
                            End If
                End Select
        Next i
    Else
'                DBGrid1(0).RowTop i   ' Light_Data(0, i)
                    DBGrid1(0).FirstRow = i
                Exit Sub
    End If
    If i >= mTotalRows&(1) Then
                MsgBox "没发现"
                
    Else
            DBGrid1(1).FirstRow = i
'            DBGrid1(1).row = i  'Check_Data(0, 10)
    End If
    
 Exit Sub

FindErr:
    Exit Sub
    mbNotFound = True

End Sub

Private Sub DBGrid1_AfterColUpdate(index As Integer, ByVal ColIndex As Integer)
 If index = 1 Then
    If ColIndex = 7 Then
        Debug.Print DBGrid1(1).Columns(3).Value, DBGrid1(1).Columns(7).Value
        Call Set_Turn(DBGrid1(1).Columns(3).Value, DBGrid1(1).Columns(7).Value)
    End If
Else
    If ColIndex = 7 Then
        Debug.Print DBGrid1(0).Columns(3).Value, DBGrid1(0).Columns(7).Value
        Call Set_Light(DBGrid1(0).Columns(3).Value, DBGrid1(0).Columns(7).Value)
    End If
End If
End Sub

Private Sub DBGrid1_BeforeColUpdate(index As Integer, ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim temp As Integer

If index = 1 Then
        If ColIndex = 7 Then
                      If Not Check_Float(DBGrid1(1).Columns(ColIndex).Value) Then
                              Call Speak_Error(DBGrid1(1).Columns(7).Caption + " 非法")
                              Cancel = True
                              DBGrid1(1).SetFocus
                              Exit Sub
                      End If
                   temp = Val(DBGrid1(1).Columns(7).Value)
                    If temp > 1 Or temp < 0 Then
                                 Call Speak_Error(DBGrid1(1).Columns(7).Caption + "=0 ,1")
                                 DBGrid1(1).SetFocus
                                Cancel = True
                                Exit Sub
                    End If
                    DBGrid1(1).Columns(7).Value = temp
        End If
End If
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

'判断按键
Private Sub DBGrid1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   Dim i, Flag As Integer
    Select Case KeyCode
        Case 40
            DBGrid1(index).SetFocus
        Case vbKeyF3
            Call Change_Command_Click
            DBGrid1(index).SetFocus
        Case vbKeyF4
            Call Print_Command_Click
            DBGrid1(index).SetFocus
        Case vbKeyF6
            Call CmdFind_Click
            DBGrid1(index).SetFocus
        Case vbKeyF10
            KeyCode = 0
    End Select
    
        
        
End Sub

'行列变化
Private Sub DBGrid1_RowColChange(index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Exit Sub
    If index = 0 Then
            DBGrid1(1).row = DBGrid1(0).row
            DBGrid1(1).col = DBGrid1(0).col
    Else
            DBGrid1(0).row = DBGrid1(1).row
            DBGrid1(0).col = DBGrid1(1).col
    End If
End Sub

'加数据
Private Sub DBGrid1_UnboundAddData(index As Integer, ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows(index) = mTotalRows(index) + 1
If index = 0 Then
    ReDim Preserve Light_Data(rec(index).Fields.count - 1, mTotalRows(index) - 1)
Else
    ReDim Preserve Check_Data(rec(index).Fields.count - 1, mTotalRows(index) - 1)
End If
NewRowBookmark = mTotalRows(index) - 1 'Sets the bookmark to the last row.
' The following loop adds a new record to the database.
For iCol = 0 To DBGrid1(index).Columns.count - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        If index = 0 Then
          Light_Data(iCol, mTotalRows(index) - 1) = RowBuf.Value(0, iCol)
        Else
          Check_Data(iCol, mTotalRows(index) - 1) = RowBuf.Value(0, iCol)
        End If
    End If
Next iCol
RowBuf.Value(0, 5) = index
RowBuf.Value(0, 6) = RowBuf.Value(0, 0)

 If index = 0 Then
    Light_Data(5, mTotalRows(index) - 1) = index ' RowBuf.Value(0, iCol)
    Light_Data(6, mTotalRows(index) - 1) = Light_Data(0, mTotalRows(index) - 1)
 Else
    Check_Data(5, mTotalRows(index) - 1) = index
    Check_Data(6, mTotalRows(index) - 1) = Check_Data(0, mTotalRows(index) - 1)
 End If


End Sub

'读数据
Private Sub DBGrid1_UnboundReadData(index As Integer, ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid is requesting rows so give them to it

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' If StartLocation is Null then start reading at the end
' or beginning of the data set.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Find the position to start reading based on the
    ' StartLocation bookmark and the iIncr variable
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfer data from our data set array to the RowBuf object
' which DBGrid uses to display the data
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows&(index) Then Exit For
        If index = 0 Then
            For iCol = 0 To UBound(Light_Data, 1)
                RowBuf.Value(iRow, iCol) = Light_Data(iCol, CurRow&)
            Next iCol
        Else
            For iCol = 0 To UBound(Check_Data, 1)
                RowBuf.Value(iRow, iCol) = Check_Data(iCol, CurRow&)
            Next iCol
        End If
    
    ' Set bookmark using CurRow& which is also our
    ' array index
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

'写数据
Private Sub DBGrid1_UnboundWriteData(index As Integer, ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Data is being updated

' Update each column in the data set array
If index = 0 Then Exit Sub
For iCol = 0 To DBGrid1(index).Columns.count - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        If index = 0 Then
            Light_Data(iCol, WriteLocation) = RowBuf.Value(0, iCol)
        Else
            Check_Data(iCol, WriteLocation) = RowBuf.Value(0, iCol)
        End If
    End If
Next iCol
End Sub

'退出
Private Sub Exit_Command_Click()
 Unload Me
End Sub
'激活窗口
Private Sub Form_Activate()
    Timer1.Interval = 900  '定时器控制
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim ShiftDown, AltDown, CtrlDown
'    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF9 Then   ' Display key combinations.
        If AltDown And CtrlDown Then
                Root_Change = Not Root_Change
                Debug.Print "F9 key"
        End If
    End If
Call Check_KeyCode(KeyCode, Shift)
End Sub

'装入窗口
Private Sub Form_Load()
Dim i, j As Integer
Width = MAIN_MDI.Width  ' Set width of form.
   Height = MAIN_MDI.Height - 900 ' Set height of form.
     Left = 0
    Top = 0
       Call Add_Win(Hwnd)
Change_Command.Tag = "N"
Set Data1(0) = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.mdb")
Set rec(0) = Data1(0).OpenRecordset(" SELECT *   from node_table where  in_out=0 order by sort_idx", dbOpenDynaset, dbInconsistent)
Set Data1(1) = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.mdb")
Set rec(1) = Data1(1).OpenRecordset(" SELECT *   from node_table where  in_out=1 order by sort_idx", dbOpenDynaset, dbInconsistent)
Root_Change = False

' Remove old columns

For j = 0 To 1
    For i = DBGrid1(j).Columns.count - 1 To 0 Step -1
        DBGrid1(j).Columns.Remove i
    Next i
    If Not rec(j).EOF() Then
            rec(j).MoveLast
    End If
    
    mTotalRows&(j) = rec(j).RecordCount
    For i = 0 To rec(j).Fields.count - 1
        DBGrid1(j).Columns.Add i
        DBGrid1(j).Columns(i).Caption = rec(j).Fields(i).Name
        DBGrid1(j).Columns(i).Visible = True
        DBGrid1(j).Columns(i).Locked = True
        DBGrid1(j).Columns(i).Width = 800
    Next i
    
    DBGrid1(j).Columns("in_out").Visible = False
    DBGrid1(j).Columns("error_name").Visible = False
    DBGrid1(j).Columns("sort_idx").Visible = False
    
        DBGrid1(j).Font.Size = 10
        DBGrid1(j).HeadFont.Size = 12
        DBGrid1(j).RowHeight = 300
        DBGrid1(j).HeadLines = 2
        
    DBGrid1(j).Columns.Add i
    DBGrid1(j).Columns(i).Visible = True
    DBGrid1(j).Columns(i).Locked = True
    DBGrid1(j).Columns(i).Width = 600
    DBGrid1(j).Columns(i).Caption = "数值"
Next j

If rec(0).RecordCount > 0 Then
    ReDim Light_Data(0 To rec(0).Fields.count, 0 To rec(0).RecordCount - 1)
    rec(0).MoveFirst
Else

    ReDim Light_Data(0 To rec(0).Fields.count, 0)
    mTotalRows&(0) = 0
End If

    
    j = 0
    Do While Not rec(0).EOF()
            For i = 0 To rec(0).Fields.count - 1
                Light_Data(i, j) = rec(0).Fields(i)
            Next i
            j = j + 1
            rec(0).MoveNext
    Loop
    
    If rec(1).RecordCount > 0 Then
        ReDim Check_Data(0 To rec(1).Fields.count, 0 To rec(1).RecordCount - 1)
        rec(1).MoveFirst
    Else
        ReDim Check_Data(0 To rec(1).Fields.count, 0)
        mTotalRows&(1) = 0
    End If
    j = 0
    Do While Not rec(1).EOF()
       
            For i = 0 To rec(1).Fields.count - 1
                Check_Data(i, j) = rec(1).Fields(i)
            Next i
            j = j + 1
            rec(1).MoveNext
    Loop
    
End Sub

'打印窗口
Private Sub Form_Paint()
Dim j As Integer

If China_English = CHINA Then
    DBGrid1(0).Caption = "运行输入信号(灯)"
    DBGrid1(1).Caption = "运行输出信号(阀)"
    If Change_Command.Tag = "N" Then
                Change_Command.Caption = "F3正常状态"""
    Else
                Change_Command.Caption = "F3调试状态"
    End If
    Exit_Command.Caption = "ESC退出"
    CmdFind.Caption = "F6查找"
    Print_Command.Caption = "F4打印"
    Exit_Command.Caption = "ESC退出"
    
    For j = 0 To 1
        DBGrid1(j).Columns(0).Caption = "描述名"

        DBGrid1(j).Columns(1).Caption = "中文名"
        DBGrid1(j).Columns(2).Caption = "英文名"
        DBGrid1(j).Columns(3).Caption = "接点名"
        DBGrid1(j).Columns(7).Caption = "数值"
    Next j
Else
    DBGrid1(0).Caption = "Running Light"
    DBGrid1(1).Caption = "Running Check"
    
    If Change_Command.Tag = "NoEdit" Then
                Change_Command.Caption = "Change"
    Else
                Change_Command.Caption = "Send"
    End If
    Exit_Command.Caption = "ESC"
    CmdFind.Caption = "F6 search"
    Print_Command.Caption = "F4 print"
    Exit_Command.Caption = "ESC exit"
    
    For j = 0 To 1
        DBGrid1(j).Columns(0).Caption = "Name"
        DBGrid1(j).Columns(1).Caption = "temp"
        DBGrid1(j).Columns(2).Caption = "china name"
        DBGrid1(j).Columns(3).Caption = "eng name"
        DBGrid1(j).Columns(4).Caption = "node name"
        DBGrid1(j).Columns(7).Caption = "Value"
    Next j
  End If

End Sub
'退出窗口
Private Sub Form_Unload(Cancel As Integer)
Dim i%
 For i% = 0 To 1
     rec(i%).Close
     Data1(i%).Close
 Next i%
Call Del_Win(Hwnd)


End Sub


'发送命令
'打印按钮
Private Sub Print_Command_Click()
   Dim Y  As Long
   Dim i%, k%, j%
   Dim temp(0 To 20)     As Single
   Dim book   As Variant
   book = DBGrid1(0).Bookmark
    
   Screen.MousePointer = vbHourglass
    If Printer.ScaleWidth < DBGrid1(0).Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        DBGrid1(0).SetFocus
        Exit Sub
    End If
    
        Printer.Height = (Elec_Input.Length + 2) * DBGrid1(0).RowHeight
        Printer.Width = Screen.Width
        temp(0) = 6
   k = 1
   With DBGrid1(0)
   For i = 0 To .Columns.count - 1
         If .Columns(i).Visible Then
                temp(k) = 0
                temp(k) = temp(k - 1) + .Columns(i).Width * (Printer.ScaleWidth) / .Width
                k = k + 1
        End If
   Next i
   

 
    
    Printer.Print
    Printer.Print
    Timer1.Enabled = False
    
    Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
    Printer.Print Trim(DBGrid1(0).Caption) + space(2) + Date$ + space(2) + time$
    
    
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
        k% = 0
        For i% = 0 To DBGrid1(0).Columns.count - 1
            If DBGrid1(0).Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(DBGrid1(0).Columns(i%).Caption); 'SetRightString(DBGRID1(0).Columns(i%).Caption, temp(K%));
                    k% = k% + 1
            End If
        Next i%
        Printer.Print String(1, " ")
        Y = Printer.CurrentY + 10
        Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
        Printer.Print String(1, " ")
    
    Dim PLine%
    j% = 0
    PLine = 0
    Do While PLine% < Elec_Input.Length
        If j% < 48 Then
            k% = 0
           For i% = 0 To DBGrid1(0).Columns.count - 1
              If DBGrid1(0).Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(Light_Data(i%, PLine)); 'SetRightString(DBGRID1(0).Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                End If
          Next i%
         Printer.Print String(1, " ")
         j% = j% + 1
         PLine = PLine + 1
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
    Timer1.Enabled = True
    'Do While j% < 48
     '       Printer.Print
      '      j% = j% + 1
    'Loop
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
   End With
    'Printer.Print space(10) + "总数:" & CStr(DCai_liao_table.Recordset.RecordCount)
'    Printer.Print "第" + CStr(Printer.Page) + "页"

    
'打印阀

'        Printer.Height = (Elec_Output.length + 2) * DBGrid1(1).RowHeight
 '       Printer.Width = Screen.Width
        temp(0) = 6
   k = 1
   With DBGrid1(1)
   For i = 0 To .Columns.count - 1
         If .Columns(i).Visible Then
                temp(k) = 0
                temp(k) = temp(k - 1) + .Columns(i).Width * (Printer.ScaleWidth) / .Width
                k = k + 1
        End If
   Next i

    Printer.Print
    Timer1.Enabled = False
    Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY)
    Printer.Print Trim(.Caption) + space(2) + Date$ + space(2) + time$
        
    Y = Printer.CurrentY + 10
    Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
    Printer.Print String(1, " ")
    
    
        k% = 0
        For i% = 0 To .Columns.count - 1
            If .Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(.Columns(i%).Caption); 'SetRightString(DBGRID1(0).Columns(i%).Caption, temp(K%));
                    k% = k% + 1
            End If
        Next i%
        Printer.Print String(1, " ")
        Y = Printer.CurrentY + 10
        Printer.Line (0, Y)-(Printer.ScaleWidth, Y) '  print --------
        Printer.Print String(1, " ")
    
    
    j% = 0
    PLine = 0
    Do While PLine% < Elec_Output.Length
        If j% < 48 Then
            k% = 0
           For i% = 0 To .Columns.count - 1
              If .Columns(i%).Visible = True Then
                    Printer.PSet (temp(k%), Printer.CurrentY)
                    Printer.Print Trim(Check_Data(i%, PLine)); 'SetRightString(DBGRID1(0).Columns(i%).Caption, temp(K%));
                    k% = k% + 1
                End If
          Next i%
         Printer.Print String(1, " ")
         j% = j% + 1
         PLine = PLine + 1
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
    
    End With

    Printer.EndDoc
    Screen.MousePointer = vbDefault
    DBGrid1(0).Bookmark = book
    Timer1.Enabled = True
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
      DBGrid1(0).Bookmark = book
    Timer1.Enabled = True
End Sub

'灯  -- elec_input
'阀  elec_output
Private Sub Timer1_Timer()
  Dim X, temp, temp_first As Long
  Dim i%, k%
  Dim get_weight, set_weight  As Single
    Dim Elec_Input      As elect_send_type
    Dim Elec_Output       As elect_send_type
   
   
 If Change_Command.Tag = "Y" And Root_Change Then
        Exit Sub
 End If
    X = read_elec_input(Elec_Input)
  If Change_Command.Tag <> "Y" Then
        X = read_elec_output(Elec_Output)
   End If
   If rec(0).RecordCount = 0 Then
            GoTo Trun_Doing
   End If
  temp_first = DBGrid1(0).FirstRow
  temp = DBGrid1(0).row
  If temp < 0 Then temp = 0
    
  For i = 0 To Elec_Input.Length
        'If Image1(i).Status <> Elec_Input.data(i) Then
                'Image1(i).Status = Elec_Input.data(i)
                '
                'DBGrid1(0).Columns(7).Value = Elec_Input.data(i)
                For k = 0 To Elec_Input.Length
                        If Elec_Input.Note_Name(i) = Light_Data(3, k) Then
                                    Light_Data(DBGrid1(0).Columns.count - 1, k) = Elec_Input.data(i)
                                    Exit For
                        End If
                Next k
        'End If
 Next i
 DBGrid1(0).Refresh
 If temp_first > 1 Then
    DBGrid1(0).FirstRow = temp_first
 End If
    DBGrid1(0).row = temp
    
Trun_Doing:
 If Change_Command.Tag = "Y" Then
        Exit Sub
 End If
 If rec(1).RecordCount = 0 Then
    Exit Sub
 End If
  temp_first = DBGrid1(1).FirstRow
  temp = DBGrid1(1).row
  If temp < 0 Then temp = 0
   
  For i = 0 To Elec_Output.Length
                For k = 0 To Elec_Output.Length
                        If Elec_Output.Note_Name(i) = Check_Data(3, k) Then
                                    Check_Data(DBGrid1(1).Columns.count - 1, k) = Elec_Output.data(i)
                                    Exit For
                        End If
                Next k
 Next i
 DBGrid1(1).Refresh
 If temp_first > 1 Then
    DBGrid1(1).FirstRow = temp_first
 End If
    DBGrid1(1).row = temp
    
End Sub

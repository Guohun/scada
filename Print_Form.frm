VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Print_Form 
   Caption         =   "报表"
   ClientHeight    =   4692
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8976
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4692
   ScaleWidth      =   8976
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\LJXT\make.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Index           =   1
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cai_Search"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\LJXT\make.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Index           =   0
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ban_search"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1092
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Print_Form.frx":0000
      Height          =   3252
      Index           =   0
      Left            =   360
      OleObjectBlob   =   "Print_Form.frx":0013
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   8252
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Print_Form.frx":1584
      Height          =   3252
      Index           =   1
      Left            =   360
      OleObjectBlob   =   "Print_Form.frx":1597
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   8252
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1104
      _Version        =   65536
      _ExtentX        =   1940
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
      Picture         =   "Print_Form.frx":2B18
   End
   Begin Threed.SSCommand Print_Command 
      Default         =   -1  'True
      Height          =   732
      Left            =   2280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
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
      MouseIcon       =   "Print_Form.frx":2F6A
      Picture         =   "Print_Form.frx":33BC
   End
End
Attribute VB_Name = "Print_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Exit_Command_Click()
Unload Me
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
Dim i, j As Integer

 Width = MAIN_MDI.Width  ' Set width of form.
 Height = MAIN_MDI.Height - 600
 Left = 0
 Top = 0
    For j = 0 To 1
        For i = 0 To DBGrid1(j).Columns.count - 1
               DBGrid1(j).Columns(i).Width = 1000
        Next i
        DBGrid1(j).Font.Size = 10
        DBGrid1(j).HeadFont.Size = 12
        DBGrid1(j).RowHeight = 500
        DBGrid1(j).HeadLines = 2
        DBGrid1(j).Refresh
  Next j
  DBGrid1(0).Columns("name").Width = 1400
  DBGrid1(0).Columns("che").Visible = False
  DBGrid1(0).Columns(7).Visible = False
  DBGrid1(1).Columns(7).Visible = False '"total_Set_weight"
  DBGrid1(1).Columns("cai_name").Width = 1200
  
  
    If zhu.Tag = "1" Then
        DBGrid1(0).Visible = True
        Call Ban_Sub
    Else
        DBGrid1(1).Visible = True
        Call Cai_Sub
    End If
    
End Sub

Private Sub Form_Paint()
    If China_English = CHINA Then
        Print_Command.Caption = "&P打印"
        Exit_Command.Caption = "ESC退出"
        DBGrid1(0).Columns("mathine").Caption = "机号"
        DBGrid1(0).Columns("ban").Caption = "班号"
        DBGrid1(0).Columns("name").Caption = "配方名"
        DBGrid1(0).Columns("m_number").Caption = "配方编号"
        DBGrid1(0).Columns("m_date").Caption = "日期"
        DBGrid1(0).Columns(5).Caption = "实际总车数"  '"mont_che"
        DBGrid1(0).Columns(7).Caption = "材料设定量"    '"total_Set_weight"
        DBGrid1(0).Columns(8).Caption = " 产量"     '"total_get_weight"
        DBGrid1(0).Columns(6).Caption = "总车数"        '"total_che"
        
        
        DBGrid1(1).Columns("name").Caption = "配方名"
        DBGrid1(1).Columns("m_number").Caption = "配方编号"
        DBGrid1(1).Columns("mathine").Caption = "机号"
        DBGrid1(1).Columns("ban").Caption = "班号"
        DBGrid1(1).Columns("cai_name").Caption = "材料名"
        DBGrid1(1).Columns("cai_number").Caption = "材料编号"
        DBGrid1(1).Columns("m_date").Caption = "日期"
        DBGrid1(1).Columns(7).Caption = "材料设定量" 'total_Set_weight
        DBGrid1(1).Columns(8).Caption = "材料消耗量" '"total_get_weight"
    Else
    End If
    
End Sub

'****************************
'班产量统计处理
'****************************
Private Sub Ban_Sub()
    Select Case zhu.Search_Text
        Case "日报表"
            Data1(0).RecordSource = "select * from ban_search"
            If Not IsEmpty(zhu.Begin_date) And zhu.Begin_date <> "" Then
                Data1(0).RecordSource = Data1(0).RecordSource & " where  m_date = #" & zhu.Begin_date & "# "
                DBGrid1(0).Columns("m_date").Visible = False
               DBGrid1(0).Caption = DBGrid1(0).Caption & zhu.Begin_date & " 日 "
            End If
            DBGrid1(0).Caption = DBGrid1(0).Caption & "  生产日报表"
    Case "月报表"
            Data1(0).RecordSource = "select 1 as m_date ,ban, mathine, m_number, name, sum(che) AS month_che, sum(total_che) AS month_total_che, sum(total_set_weight) AS month_set, sum(total_get_weight) AS month_get " _
                                                & " From ban_search  "
           If Not IsEmpty(zhu.End_date) And zhu.End_date <> "" Then
                 Data1(0).RecordSource = Data1(0).RecordSource & " where  Year(m_date)= " & zhu.End_date
                 DBGrid1(0).Columns("m_date").Visible = False  '""
                 DBGrid1(0).Caption = DBGrid1(0).Caption & zhu.End_date & " 年 "
            End If
            If Not IsEmpty(zhu.Begin_date) And zhu.Begin_date <> "" Then
               If Right(Data1(0).RecordSource, 10) <> "ban_search" Then
                        Data1(0).RecordSource = Data1(0).RecordSource & "  and  Month(m_date)= " & zhu.Begin_date
                        DBGrid1(0).Columns("m_date").Visible = False  '""
                        DBGrid1(0).Caption = DBGrid1(0).Caption & zhu.Begin_date & " 月 "
               Else
                        Data1(0).RecordSource = Data1(0).RecordSource & " where  Month(m_date)= " & zhu.Begin_date
               End If
            End If
            DBGrid1(0).Caption = DBGrid1(0).Caption & "  生产月报表"
    Case "指定报表"
            Data1(0).RecordSource = "select 1 as m_date ,ban, mathine, m_number, name, sum(che) AS month_che, sum(total_che) AS month_total_che, sum(total_set_weight) AS month_set, sum(total_get_weight) AS month_get " _
                                                & " From ban_search  "
            If Not IsEmpty(zhu.Begin_date) And zhu.Begin_date <> "" Then
                Data1(0).RecordSource = Data1(0).RecordSource & " where  m_date>=#" & zhu.Begin_date & "# and m_date<=#" & zhu.End_date & "#"
                DBGrid1(0).Columns("m_date").Visible = False  '""
                 DBGrid1(0).Caption = DBGrid1(0).Caption & zhu.Begin_date & " ~ " & zhu.End_date
            End If
            DBGrid1(0).Caption = DBGrid1(0).Caption & "  生产指定日期报表"

    End Select
            If zhu.Ban_hao <> 0 Then
               If Right(Data1(0).RecordSource, 10) <> "ban_search" Then
                   Data1(0).RecordSource = Data1(0).RecordSource & "  and ban = " & zhu.Ban_hao
                Else
                  Data1(0).RecordSource = Data1(0).RecordSource & " where  ban = " & zhu.Ban_hao
                End If
            Else
                DBGrid1(0).Columns("ban").Visible = False
            End If
            If zhu.Search_Mathine <> 0 Then
                zhu.Search_Mathine = zhu.Search_Mathine Mod 3
               If Right(Data1(0).RecordSource, 10) <> "ban_search" Then
                Data1(0).RecordSource = Data1(0).RecordSource & "  and mathine = " & zhu.Search_Mathine
                Else
                Data1(0).RecordSource = Data1(0).RecordSource & " where mathine = " & zhu.Search_Mathine
                End If
            Else
                DBGrid1(0).Columns("mathine").Visible = False
            End If
            If Not IsEmpty(zhu.Search_Number) And zhu.Search_Number <> "" Then
               If Right(Data1(0).RecordSource, 10) <> "ban_search" Then
                Data1(0).RecordSource = Data1(0).RecordSource & "  and m_number like '" & zhu.Search_Number & "'"
                Else
                Data1(0).RecordSource = Data1(0).RecordSource & " where m_number like  '" & zhu.Search_Number & "'"
                End If
                DBGrid1(0).Caption = DBGrid1(0).Caption & zhu.Search_Number & "配方 "
            Else
                'DBGrid1(0).Columns("m_number").Visible = False
                'DBGrid1(0).Columns("name").Visible = False
            End If
    Select Case zhu.Search_Text
    Case "日报表"
                DBGrid1(0).Columns(6).DataField = "total_che"
                DBGrid1(0).Columns(7).DataField = "total_set_weight"
                DBGrid1(0).Columns(8).DataField = "total_get_weight"
        Case "月报表", "指定报表"
                Data1(0).RecordSource = Data1(0).RecordSource & " group by ban, mathine, m_number, name"
                DBGrid1(0).Columns(6).DataField = "month_total_che"
                DBGrid1(0).Columns(7).DataField = "month_set"
                DBGrid1(0).Columns(8).DataField = "month_get"
    End Select
            Data1(0).Refresh
            DBGrid1(0).Refresh

End Sub

'****************************
'材料产量统计处理
'****************************
Private Sub Cai_Sub()
    Select Case zhu.Search_Text
        Case "日报表"
            If Not IsEmpty(zhu.Search_Number) And zhu.Search_Number <> "" Then
                    Data1(1).RecordSource = " select * from cai_search"
                Else
                    Data1(1).RecordSource = " select m_date,ban,mathine,1 as m_number,2  as name,cai_number,cai_name,sum(total_set_weight) as set_weight ,sum(total_get_weight) as  get_weight  from cai_search"
            End If
            If Not IsEmpty(zhu.Begin_date) And zhu.Begin_date <> "" Then
                Data1(1).RecordSource = Data1(1).RecordSource & " where  m_date = #" & zhu.Begin_date & "# "
                DBGrid1(1).Columns("m_date").Visible = False
                DBGrid1(1).Caption = zhu.Begin_date + space(1)
            End If
            DBGrid1(1).Caption = DBGrid1(1).Caption & "  材料消耗日报表"
        Case "月报表"
            Data1(1).RecordSource = "SELECT 1 as m_date,ban,mathine,2 as m_number," _
                    & " 3 as name,cai_number,cai_name,sum (total_set_weight) as month_set ," _
                    & " sum(total_get_weight) as month_get from cai_search   "
           If Not IsEmpty(zhu.End_date) And zhu.End_date <> "" Then
                 Data1(1).RecordSource = Data1(1).RecordSource & " where  Year(m_date)= " & zhu.End_date
                 DBGrid1(1).Columns("m_date").Visible = False  '""
                 DBGrid1(1).Caption = DBGrid1(1).Caption & zhu.End_date & " 年 "
            End If
            If Not IsEmpty(zhu.Begin_date) And zhu.Begin_date <> "" Then
            If Right(Data1(1).RecordSource, 10) <> "cai_search" Then
                Data1(1).RecordSource = Data1(1).RecordSource & "  and month(m_date) =" & zhu.Begin_date
            Else
                Data1(1).RecordSource = Data1(1).RecordSource & " where  month(m_date) =" & zhu.Begin_date
            End If
                DBGrid1(1).Columns("m_date").Visible = False
                DBGrid1(1).Caption = DBGrid1(1).Caption & zhu.Begin_date & " 月 "
            End If
            DBGrid1(1).Caption = DBGrid1(1).Caption & "  材料消耗月报表"
    Case "指定报表"
            Data1(1).RecordSource = "SELECT 1 as m_date,ban,mathine,2 as m_number," _
                    & " 3 as name,cai_number,cai_name,sum (total_set_weight) as month_set ," _
                    & " sum(total_get_weight) as month_get from cai_search   "
            If Not IsEmpty(zhu.Begin_date) And zhu.Begin_date <> "" Then
                Data1(1).RecordSource = Data1(1).RecordSource & " where  m_date>=#" & zhu.Begin_date & "# and m_date<=#" & zhu.End_date & "#"
                DBGrid1(1).Columns("m_date").Visible = False
                DBGrid1(1).Caption = DBGrid1(1).Caption & zhu.Begin_date & " ~ " & zhu.End_date
            End If
            DBGrid1(1).Caption = DBGrid1(1).Caption & "  材料消耗指定日报表"
    End Select
    
            If zhu.Ban_hao <> 0 Then
               If Right(Data1(1).RecordSource, 10) <> "cai_search" Then
                Data1(1).RecordSource = Data1(1).RecordSource & "  and ban = " & zhu.Ban_hao
                Else
                Data1(1).RecordSource = Data1(1).RecordSource & " where  ban = " & zhu.Ban_hao
                End If
            Else
                DBGrid1(1).Columns("ban").Visible = False
            End If
            If zhu.Search_Mathine <> 0 Then
               zhu.Search_Mathine = zhu.Search_Mathine Mod 3
               If Right(Data1(1).RecordSource, 10) <> "cai_search" Then
                Data1(1).RecordSource = Data1(1).RecordSource & "  and mathine = " & zhu.Search_Mathine
                Else
                Data1(1).RecordSource = Data1(1).RecordSource & " where mathine = " & zhu.Search_Mathine
                End If
                
            Else
                DBGrid1(1).Columns("mathine").Visible = False
            End If
            If Not IsEmpty(zhu.Search_Number) And zhu.Search_Number <> "" Then
               If Right(Data1(1).RecordSource, 10) <> "cai_search" Then
                Data1(1).RecordSource = Data1(1).RecordSource & "  and m_number like '" & zhu.Search_Number & "'"
                Else
                Data1(1).RecordSource = Data1(1).RecordSource & " where m_number like  '" & zhu.Search_Number & "'"
                End If
                DBGrid1(1).Caption = DBGrid1(1).Caption & zhu.Search_Number & "配方 "
            End If
            DBGrid1(1).Columns("m_number").Visible = False
             DBGrid1(1).Columns("name").Visible = False
            
    Select Case zhu.Search_Text
    Case "日报表"
                If Not IsEmpty(zhu.Search_Number) And zhu.Search_Number <> "" Then
                    'Data1(1).RecordSource = "select * from cai_search"
                Else
                    Data1(1).RecordSource = Data1(1).RecordSource & " group by m_date,ban,mathine,cai_number,cai_name"
                    DBGrid1(1).Columns(7).DataField = "set_weight"
                    DBGrid1(1).Columns(8).DataField = "get_weight"
                    
                End If
                Debug.Print Data1(1).RecordSource
                
        Case "月报表", "指定报表"
                Data1(1).RecordSource = Data1(1).RecordSource & " group by ban,mathine,cai_number,cai_name  "
                DBGrid1(1).Columns(7).DataField = "month_set"
                DBGrid1(1).Columns(8).DataField = "month_get"
    End Select
           Debug.Print Data1(1).RecordSource
           
            Data1(1).Refresh
            DBGrid1(1).Refresh

End Sub

Private Sub Print_Command_Click()
   'On Error GoTo error_doing
   Dim Y  As Long
   Dim i%, k%
   Dim temp(0 To 20)     As Long
   Dim book As Variant
   Dim index%
If zhu.Tag = "1" Then
    index = 0
Else
    index = 1
End If
   book = DBGrid1(index).Bookmark
   
   Screen.MousePointer = vbHourglass
    If Printer.ScaleWidth < DBGrid1(index).Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        DBGrid1(index).SetFocus
        Exit Sub
    End If
  With Data1(index)
        Printer.Height = (.Recordset.RecordCount + 2) * DBGrid1(index).RowHeight
        Printer.Width = Screen.Width
   End With
    
    
   
   
   
   temp(0) = 6
   k = 1
   
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

        Printer.PSet (Printer.ScaleWidth / 3, Printer.CurrentY), 2
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
    Data1(index).Refresh
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
   
'    Printer.Print space(10) + "总数:" & CStr(DCai_liao_table.Recordset.RecordCount)
 '   Printer.Print "第" + CStr(Printer.Page) + "页"
    Printer.EndDoc
    
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
    DBGrid1(index).SetFocus

End Sub

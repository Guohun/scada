VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Produce_Form 
   Caption         =   "Form1"
   ClientHeight    =   5568
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8736
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5568
   ScaleWidth      =   8736
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Produce_Form.frx":0000
      Height          =   1332
      Index           =   0
      Left            =   480
      OleObjectBlob   =   "Produce_Form.frx":0013
      TabIndex        =   0
      Top             =   2760
      Width           =   7932
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Produce_Form.frx":1A7C
      Height          =   852
      Index           =   1
      Left            =   480
      OleObjectBlob   =   "Produce_Form.frx":1A8F
      TabIndex        =   1
      Top             =   4440
      Width           =   7932
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Produce_Form.frx":3194
      Height          =   1812
      Index           =   2
      Left            =   480
      OleObjectBlob   =   "Produce_Form.frx":31A7
      TabIndex        =   2
      Top             =   720
      Width           =   7932
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   5040
      TabIndex        =   3
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
      Picture         =   "Produce_Form.frx":4A74
   End
   Begin Threed.SSCommand Print_Command 
      Height          =   732
      Left            =   2760
      TabIndex        =   4
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
      MouseIcon       =   "Produce_Form.frx":4EC6
      Picture         =   "Produce_Form.frx":5318
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\LJXT\make.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Index           =   2
      Left            =   4440
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Che_Search"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\LJXT\make.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Index           =   1
      Left            =   2760
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Base_Search"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1452
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
      Left            =   960
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Duan_Search"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1452
   End
End
Attribute VB_Name = "Produce_Form"
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

 Width = MAIN_MDI.Width  ' Set width of form.
 Height = MAIN_MDI.Height - 600
 Left = 0
 Top = 0
 Call Add_Win(Hwnd)
 
 Call Duan_Search
 Call Che_Search
Call Base_Search

End Sub

Private Sub Form_Paint()
If China_English = CHINA Then
        Caption = "混练过程显示"
        DBGrid1(0).Columns("eng_str").Visible = False
        DBGrid1(0).Columns("china_str").Visible = True
        Print_Command.Caption = "&P打印"
        Exit_Command.Caption = "ESC退出"

        DBGrid1(0).Columns("mathine").Caption = "机号"
        DBGrid1(0).Columns("ban").Caption = "班号"
        DBGrid1(0).Columns("name").Caption = "配方名"
        DBGrid1(0).Columns("m_number").Caption = "配方编号"
        DBGrid1(0).Columns("m_date").Caption = "日期"
        DBGrid1(0).Columns("duan_hao").Caption = "段号"
        DBGrid1(0).Columns("china_str").Caption = "功能"
        DBGrid1(0).Columns("get_the_time").Caption = "时间"
        DBGrid1(0).Columns("next_power").Caption = "能量"
        DBGrid1(0).Columns("next_tempro").Caption = "温度"
        DBGrid1(0).Columns("set_speed").Caption = "转速"
        
        DBGrid1(1).Columns("name").Caption = "配方名"
        DBGrid1(1).Columns("m_number").Caption = "配方编号"
        DBGrid1(1).Columns("mathine").Caption = "机号"
        DBGrid1(1).Columns("ban").Caption = "班号"
        DBGrid1(1).Columns("m_date").Caption = "日期"
        DBGrid1(1).Columns("total_time").Caption = "排料时间"
        DBGrid1(1).Columns("all_power").Caption = "总能量"
        DBGrid1(1).Columns("pai_tempro").Caption = "排料温度"
        DBGrid1(1).Columns("huan_time").Caption = "混炼总时间"
        
        DBGrid1(2).Columns("cai_name").Caption = "材料名"
        DBGrid1(2).Columns("cai_number").Caption = "材料编号"
        DBGrid1(2).Columns("total_set_weight").Caption = "设定量"
        DBGrid1(2).Columns("total_get_weight").Caption = "消耗量"
        
'        DBGrid1(2).Columns("w1").Caption = "材料设定量"
 '       DBGrid1(2).Columns("w2").Caption = "材料消耗量"
    Else
           DBGrid1(0).Columns("eng_str").Visible = True
            DBGrid1(0).Columns("china_str").Visible = False

    End If
    
End Sub


'****************************
'段统计表处理
'****************************
Private Sub Duan_Search()
            Data1(0).RecordSource = "select * from duan_search"
            If Not IsEmpty(zhu_pro.Begin_date) And zhu_pro.Begin_date <> "" Then
                Data1(0).RecordSource = Data1(0).RecordSource & " where  m_date = #" & zhu_pro.Begin_date & "# "
                DBGrid1(0).Caption = zhu_pro.Begin_date + space(1)
            End If
            If Not IsNull(zhu_pro.Ban_hao) Then
               If Right(Data1(0).RecordSource, 11) <> "duan_search" Then
                Data1(0).RecordSource = Data1(0).RecordSource & "  and ban = " & zhu_pro.Ban_hao
                Else
                Data1(0).RecordSource = Data1(0).RecordSource & " where  ban = " & zhu_pro.Ban_hao
                End If
                
            Else
                'DBGrid1(0).Columns("ban").Visible = False
            End If
            If Not IsNull(zhu_pro.Search_Mathine) Then
               If Right(Data1(0).RecordSource, 11) <> "duan_search" Then
                Data1(0).RecordSource = Data1(0).RecordSource & "  and mathine = " & zhu_pro.Search_Mathine
                Else
                Data1(0).RecordSource = Data1(0).RecordSource & " where mathine = " & zhu_pro.Search_Mathine
                End If
                DBGrid1(0).Caption = DBGrid1(0).Caption + CStr(zhu_pro.Search_Mathine) + "号机组 "
            Else
                'DBGrid1(0).Columns("mathine").Visible = False
            End If
            If Not IsEmpty(zhu_pro.Search_Number) And zhu_pro.Search_Number <> "" Then
               If Right(Data1(0).RecordSource, 11) <> "duan_search" Then
                Data1(0).RecordSource = Data1(0).RecordSource & "  and m_number like '" & zhu_pro.Search_Number & "'"
                Else
                Data1(0).RecordSource = Data1(0).RecordSource & " where m_number like  '" & zhu_pro.Search_Number & "'"
                End If
                DBGrid1(0).Caption = DBGrid1(0).Caption & zhu_pro.Search_Number & "配方 "
            Else
                'DBGrid1(0).Columns("m_number").Visible = False
                'DBGrid1(0).Columns("name").Visible = False
            End If
            If Not IsNull(zhu_pro.Now_Che) Then
               If Right(Data1(0).RecordSource, 11) <> "duan_search" Then
                Data1(0).RecordSource = Data1(0).RecordSource & "  and now_che = " & zhu_pro.Now_Che & "  "
                Else
                Data1(0).RecordSource = Data1(0).RecordSource & "  where  now_che = " & zhu_pro.Now_Che & " "
                End If
                DBGrid1(0).Caption = DBGrid1(0).Caption + str(zhu_pro.Now_Che) + "车"
            Else
            End If
            Debug.Print Data1(0).RecordSource
            Data1(0).RecordSource = Data1(0).RecordSource & " order  by duan_hao"
            DBGrid1(0).Caption = DBGrid1(0).Caption & "  各段功能表"
            Data1(0).Refresh
            DBGrid1(0).Refresh
            
End Sub
'****************************
'密炼过程统计表处理
'****************************
Private Sub Che_Search()
            Data1(2).RecordSource = "select * from che_search"
            If Not IsEmpty(zhu_pro.Begin_date) And zhu_pro.Begin_date <> "" Then
                Data1(2).RecordSource = Data1(2).RecordSource & " where  m_date = #" & zhu_pro.Begin_date & "#   "
                DBGrid1(2).Caption = zhu_pro.Begin_date + space(1)
                'DBGrid1(0).Columns("m_date").Visible = False
            End If
            If Not IsNull(zhu_pro.Ban_hao) Then
               If Right(Data1(2).RecordSource, 1) = " " Then
                Data1(2).RecordSource = Data1(2).RecordSource & "  and ban = " & zhu_pro.Ban_hao & "  "
                Else
                Data1(2).RecordSource = Data1(2).RecordSource & " where  ban = " & zhu_pro.Ban_hao & " "
                End If
            Else
                'DBGrid1(0).Columns("ban").Visible = False
            End If
            If Not IsNull(zhu_pro.Search_Mathine) Then
               If Right(Data1(2).RecordSource, 1) = " " Then
                Data1(2).RecordSource = Data1(2).RecordSource & "  and mathine = " & zhu_pro.Search_Mathine & " "
                Else
                Data1(2).RecordSource = Data1(2).RecordSource & " where mathine = " & zhu_pro.Search_Mathine & "  "
                End If
                DBGrid1(2).Caption = DBGrid1(2).Caption + CStr(zhu_pro.Search_Mathine) + "号机组 "
            Else
                'DBGrid1(0).Columns("mathine").Visible = False
            End If
            If Not IsEmpty(zhu_pro.Search_Number) And zhu_pro.Search_Number <> "" Then
               If Right(Data1(2).RecordSource, 1) = " " Then
                Data1(2).RecordSource = Data1(2).RecordSource & "  and m_number like '" & zhu_pro.Search_Number & "' "
                Else
                Data1(2).RecordSource = Data1(2).RecordSource & " where m_number like  '" & zhu_pro.Search_Number & "' "
                End If
                DBGrid1(2).Caption = DBGrid1(2).Caption & zhu_pro.Search_Number & "配方 "
            Else
                'DBGrid1(0).Columns("m_number").Visible = False
                'DBGrid1(0).Columns("name").Visible = False
            End If
            If Not IsNull(zhu_pro.Now_Che) Then
               If Right(Data1(2).RecordSource, 1) = " " Then
                Data1(2).RecordSource = Data1(2).RecordSource & "  and now_che = " & zhu_pro.Now_Che & "  "
                Else
                Data1(2).RecordSource = Data1(2).RecordSource & " where  now_che = " & zhu_pro.Now_Che & " "
                End If
                DBGrid1(2).Caption = DBGrid1(2).Caption + str(zhu_pro.Now_Che) + "车"
            Else
            End If
        
            Debug.Print Data1(2).RecordSource
            
            DBGrid1(2).Caption = DBGrid1(2).Caption & "  生产材料消耗表"
            Data1(2).Refresh
            DBGrid1(2).Refresh


End Sub

'****************************
'基本统计表处理
'****************************
Private Sub Base_Search()
            Data1(1).RecordSource = "select * from base_search"
            If Not IsEmpty(zhu_pro.Begin_date) And zhu_pro.Begin_date <> "" Then
                Data1(1).RecordSource = Data1(1).RecordSource & " where  m_date = #" & zhu_pro.Begin_date & "#   "
                'DBGrid1(0).Columns("m_date").Visible = False
                DBGrid1(1).Caption = zhu_pro.Begin_date + space(1)
            End If
                        
            If Not IsNull(zhu_pro.Ban_hao) Then
               If Right(Data1(1).RecordSource, 1) = " " Then
                Data1(1).RecordSource = Data1(1).RecordSource & "  and ban = " & zhu_pro.Ban_hao & "  "
                Else
                Data1(1).RecordSource = Data1(1).RecordSource & " where  ban = " & zhu_pro.Ban_hao & " "
                End If
            Else
                'DBGrid1(0).Columns("ban").Visible = False
            End If
            If Not IsNull(zhu_pro.Search_Mathine) Then
               If Right(Data1(1).RecordSource, 1) = " " Then
                Data1(1).RecordSource = Data1(1).RecordSource & "  and mathine = " & zhu_pro.Search_Mathine & " "
                Else
                Data1(1).RecordSource = Data1(1).RecordSource & " where mathine = " & zhu_pro.Search_Mathine & "  "
                End If
                DBGrid1(1).Caption = DBGrid1(1).Caption + CStr(zhu_pro.Search_Mathine) + "号机组 "
            Else
                'DBGrid1(0).Columns("mathine").Visible = False
            End If
            If Not IsEmpty(zhu_pro.Search_Number) And zhu_pro.Search_Number <> "" Then
               If Right(Data1(1).RecordSource, 1) = " " Then
                Data1(1).RecordSource = Data1(1).RecordSource & "  and m_number like '" & zhu_pro.Search_Number & "' "
                Else
                Data1(1).RecordSource = Data1(1).RecordSource & " where m_number like  '" & zhu_pro.Search_Number & "' "
                End If
                DBGrid1(1).Caption = DBGrid1(1).Caption & zhu_pro.Search_Number & "配方 "
            Else
                'DBGrid1(0).Columns("m_number").Visible = False
                'DBGrid1(0).Columns("name").Visible = False
            End If
            If Not IsNull(zhu_pro.Now_Che) Then
               If Right(Data1(1).RecordSource, 1) = " " Then
                Data1(1).RecordSource = Data1(1).RecordSource & "  and now_che = " & zhu_pro.Now_Che & "  "
                Else
                Data1(1).RecordSource = Data1(1).RecordSource & " where  now_che = " & zhu_pro.Now_Che & " "
                End If
                DBGrid1(1).Caption = DBGrid1(1).Caption + str(zhu_pro.Now_Che) + "车"
            Else
            End If
            
            Debug.Print Data1(1).RecordSource
            DBGrid1(1).Caption = DBGrid1(1).Caption & "  生产时间统计表"
            Data1(1).Refresh
            
            'DBGrid1(1).Refresh

End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call Del_Win(Me.Hwnd)
End Sub

'use max dbgrid to print
'
Private Sub Print_Command_Click()
   Dim Y  As Long
   Dim i%, k%
   Dim temp(0 To 20)     As Single
    Dim Print_time%
   Dim index%
   Dim My_Index
   My_Index = Array(2, 0, 1)
   Screen.MousePointer = vbHourglass
    If Printer.ScaleWidth < DBGrid1(0).Width Then
        Speak_Error ("打印机宽度太小，请从控制面板设置打印机")
        Screen.MousePointer = vbDefault
        DBGrid1(0).SetFocus
        Exit Sub
    End If
    
    Dim Pape_Len%
    Pape_Len = 0
    For Print_time% = 0 To 2
    With Data1(Print_time%)
         Pape_Len = Pape_Len + (.Recordset.RecordCount + 2) * DBGrid1(Print_time%).RowHeight
     End With
   Next Print_time%
   Printer.Height = Pape_Len
   Printer.Width = Screen.Width
    
For Print_time% = 0 To 2
    index = My_Index(Print_time)
   
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
    
    Next Print_time
    
    Printer.EndDoc
    Erase My_Index
    Screen.MousePointer = vbDefault
    
    Exit Sub
error_doing:
    MsgBox "打印机有问题"
    Screen.MousePointer = vbDefault
    DBGrid1(2).SetFocus
    Erase My_Index

End Sub

VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form Power_Form 
   Caption         =   "能量"
   ClientHeight    =   6792
   ClientLeft      =   948
   ClientTop       =   1740
   ClientWidth     =   9228
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6792
   ScaleWidth      =   9228
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9225
      _Version        =   65536
      _ExtentX        =   16272
      _ExtentY        =   1720
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand Exit_Command 
         Cancel          =   -1  'True
         Height          =   732
         Left            =   6312
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
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
         Picture         =   "Zu_gp.frx":0000
      End
      Begin Threed.SSCommand Print_Command 
         Height          =   732
         Left            =   2280
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
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
         MouseIcon       =   "Zu_gp.frx":0452
         Picture         =   "Zu_gp.frx":08A4
      End
      Begin Threed.SSCommand Power_eng 
         Default         =   -1  'True
         Height          =   732
         Left            =   4296
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "power"
         Top             =   120
         Width           =   1104
         _Version        =   65536
         _ExtentX        =   1940
         _ExtentY        =   1291
         _StockProps     =   78
         Caption         =   "&Power"
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
         MouseIcon       =   "Zu_gp.frx":0CF6
         Picture         =   "Zu_gp.frx":1148
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   5280
      Left            =   0
      TabIndex        =   0
      Top             =   972
      Width           =   9228
      _Version        =   65536
      _ExtentX        =   16277
      _ExtentY        =   9313
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   2
      Begin MSChartLib.MSChart Power_Graph 
         Height          =   3732
         Left            =   360
         OleObjectBlob   =   "Zu_gp.frx":159A
         TabIndex        =   5
         Top             =   360
         Width           =   8052
      End
   End
End
Attribute VB_Name = "Power_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim Max_Point As Integer


'读能量数据
Private Function Power_Read() As String
Dim db As Database, rs  As Recordset
Dim sqlstr  As String
 Dim Temp_Value   As Single
 
Set db = OpenDatabase(Data_Path & "\make.mdb")
    sqlstr = "select  power  from grap_search"
    If Not IsEmpty(zhu_pro.Begin_date) And zhu_pro.Begin_date <> "" Then
        sqlstr = sqlstr & " where  m_date = #" & zhu_pro.Begin_date & "#   "
    End If
    If Not IsNull(zhu_pro.Ban_hao) Then
       If Right(sqlstr, 1) = " " Then
        sqlstr = sqlstr & "  and ban = " & zhu_pro.Ban_hao & "  "
        Else
        sqlstr = sqlstr & " where  ban = " & zhu_pro.Ban_hao & " "
        End If
    Else
    End If
    If Not IsNull(zhu_pro.Search_Mathine) Then
       If Right(sqlstr, 1) = " " Then
        sqlstr = sqlstr & "  and mathine = " & zhu_pro.Search_Mathine & " "
        Else
        sqlstr = sqlstr & " where mathine = " & zhu_pro.Search_Mathine & "  "
        End If
    Else
        'DBGrid1(0).Columns("mathine").Visible = False
    End If
    If Not IsEmpty(zhu_pro.Search_Number) And zhu_pro.Search_Number <> "" Then
       If Right(sqlstr, 1) = " " Then
        sqlstr = sqlstr & "  and m_number like '" & zhu_pro.Search_Number & "' "
        Else
        sqlstr = sqlstr & " where m_number like  '" & zhu_pro.Search_Number & "' "
        End If
    Else
    End If
    If Not IsNull(zhu_pro.Now_Che) Then
                sqlstr = sqlstr & "  and now_che = " & zhu_pro.Now_Che & "  "
                Else
                sqlstr = sqlstr & "  where  now_che = " & zhu_pro.Now_Che & " "
    End If
                           
    
    Set rs = db.OpenRecordset(sqlstr)
    Power_Read = ""
    temp_vlaue = 0
    With rs
        Do While Not .EOF
                        If (Power_Read = "") Then
                                Power_Read = vFieldVal(.Fields("power").Value)
                        Else
                                Power_Read = Power_Read & "," & vFieldVal(.Fields("power").Value)
                        End If
                        Debug.Print , Temp_Value;
                .MoveNext
        Loop
 End With
rs.Close
db.Close
End Function


Private Sub Exit_Command_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
Dim Temp_Str    As String
Max_Point = 3600
Width = MAIN_MDI.Width  ' Set width of form.
Height = MAIN_MDI.Height * 3 / 4 ' Set height of form.
Left = 0
Top = 0
Me.WindowState = 2
  Power_eng.Tag = "eng"
  
Call Add_Win(Hwnd)
Temp_Str = Power_Read()
Power_Graph.Visible = False
Call Send_Data_To_Graph(Power_Graph, 0, Temp_Str)
Power_Graph.Visible = True
End Sub

Private Sub Form_Paint()
Dim axisID
    Call First_Do
    If China_English = CHINA Then
        Exit_Command.Caption = "ESC退出"
         Print_Command.Caption = "&P图形输出"
         If Power_eng.Tag = "power" Then
                Power_eng.Caption = "&Z功率"
        Else
                Power_eng.Caption = "&Z能量"
        End If
        For axisID = VtChAxisIdX To VtChAxisIdZ
        With Power_Graph.Plot.Axis(axisID, 1).AxisTitle
            .Visible = True
            Select Case axisID
                Case 0
                    .Text = "时间 （秒)"
                Case 1
                    If Power_eng.Tag = "power" Then
                        .Text = "功率"
                    Else
                        .Text = "能量"
                    End If
                Case 2
                    .Text = ""
                Case 3
                    .Text = ""

                End Select
        End With
        Next
    Else
        Exit_Command.Caption = "Exit"
        Print_Command.Caption = "&Print Grap"

         If Power_eng.Tag = "power" Then
                Power_eng.Caption = "&ZPower->Eng"
        Else
                Power_eng.Caption = "&ZEng->Power"
        End If
        For axisID = VtChAxisIdX To VtChAxisIdZ
        With Power_Graph.Plot.Axis(axisID, 1).AxisTitle
            .Visible = True
            Select Case axisID
                Case 0
                    .Text = "Timer (s)"
                Case 1
                    If Power_eng.Tag = "power" Then
                        .Text = "power"
                    Else
                        .Text = "eng"
                    End If
                Case 2
                    .Text = ""
                Case 3
                    .Text = ""

                End Select
        End With
        Next
        
    End If
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call Del_Win(Hwnd)
End Sub

Private Sub First_Do()
With Power_Graph
  .Left = 0: .Top = 0
  .Width = SSPanel1.Width
  .Height = SSPanel1.Height
End With
If China_English = CHINA Then
    Caption = "能量曲线图"
    'Power_Graph.Title(CHART_TOPTIT) = zhu_pro.Begin_date + space(3) + zhu_pro.Search_Number + "配方 第" + str(zhu_pro.Now_Che) + "车能量曲线图"
Else
    Caption = "The Graphics of  Power"
   'Power_Graph.Title(CHART_TOPTIT) = zhu_pro.Begin_date + space(3) + zhu_pro.Search_Number + "   " + str(zhu_pro.Now_Che) + "The Graphics of  Power"
End If
End Sub



Private Sub Power_eng_Click()
        If Power_eng.Tag = "power" Then
            Power_eng.Tag = "eng"
        Else
            Power_eng.Tag = "power"
        End If

        Temp_Str = Power_Read()
        Power_Graph.Visible = False
        Call Send_Data_To_Graph(Power_Graph, 0, Temp_Str)
        Power_Graph.Visible = True
        Me.Refresh
End Sub

Private Sub Print_Command_Click()
Dim p As Printer
Set p = Printer
With Power_Graph
'    .Style = CS_PRINTABLE
'    .Paint ' p.hdc, 0, 0, 1500, 1500, True, 0
  '  .PrintIt
End With
End Sub



Private Function Send_Data_To_Graph(C As MSChart, Line As Integer, ByVal s As String)
Dim i, Last_Pos, Pos As Integer
Dim Val As Integer
i = 1:  Last_Pos = 0:   Pos = InStr(1, s, ",", vbTextCompare)
Val = 0
C.Visible = False
With C
    While ((Pos <> 0) Or (Pos <> Null))
                If Power_eng.Tag = "eng" Then
                    Val = Val + CInt(Mid(s, Last_Pos + 1, (Pos - 1) - (Last_Pos + 1) + 1))
                Else
                    Val = CInt(Mid(s, Last_Pos + 1, (Pos - 1) - (Last_Pos + 1) + 1))
                End If
                .RowCount = i
                .row = i
                .RowLabel = Vround((i - 1) * Time_Simpling, 1)
                .data = Val
                
                i = i + 1
            
        Last_Pos = Pos:        Pos = InStr(Pos + 1, s, ",", vbTextCompare)
    Wend
End With
C.Visible = True
End Function


VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form bl 
   Caption         =   "Form2"
   ClientHeight    =   6540
   ClientLeft      =   768
   ClientTop       =   2652
   ClientWidth     =   9468
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6540
   ScaleWidth      =   9468
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   972
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9468
      _Version        =   65536
      _ExtentX        =   16700
      _ExtentY        =   1714
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
         Left            =   4920
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
         Picture         =   "Zu_gph.frx":0000
      End
      Begin Threed.SSCommand Print_Command 
         Default         =   -1  'True
         Height          =   732
         Left            =   2160
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
         MouseIcon       =   "Zu_gph.frx":0452
         Picture         =   "Zu_gph.frx":08A4
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   5640
      Left            =   0
      TabIndex        =   1
      Top             =   972
      Width           =   9468
      _Version        =   65536
      _ExtentX        =   16700
      _ExtentY        =   9948
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
      Alignment       =   4
      Autosize        =   1
      Begin MSChartLib.MSChart Tempro_Graph 
         Height          =   3972
         Left            =   480
         OleObjectBlob   =   "Zu_gph.frx":0CF6
         TabIndex        =   4
         Top             =   360
         Width           =   8412
      End
   End
End
Attribute VB_Name = "bl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Const CD_ALL = 0
Const CD_VALUES = 1
Const CD_YLEG = 2
Const CD_XLEG = 3

Const CSA_MIN = 0
Const CSA_MAX = 1
Const CSA_GAP = 2
Const CSA_SCALE = 3
Const CSA_YLEGGAP = 4
Const CSA_PIXXVALUE = 5
Const CSA_XMIN = 6
Const CSA_XMAX = 7
Const CSA_XGAP = 8

Const CL_NOTCLIPPED = 1
Const CL_NOTCHANGECOLOR = 2
Const CL_HIDE = 4
Const CL_FORCESERLEG = 8
Const CL_GETLEGEND = 16
Const CC_HIDETEXT = 1
Const CC_HIDE = 2

Dim Max_Point As Integer
Private Function Send_Data_To_Graph(C As MSChart, Line As Integer, ByVal s As String)
Dim i, Last_Pos, Pos As Integer
Dim Val As Integer
i = 1:  Last_Pos = 0:   Pos = InStr(1, s, ",", vbTextCompare)
k = 1
C.Visible = False
With C
    While ((Pos <> 0) Or (Pos <> Null))
                .RowCount = i
                .row = i
                .data = CInt(Mid(s, Last_Pos + 1, (Pos - 1) - (Last_Pos + 1) + 1))
                .RowLabel = Vround((i - 1) * Time_Simpling, 1)
                i = i + 1
        Last_Pos = Pos:        Pos = InStr(Pos + 1, s, ",", vbTextCompare)
    Wend
End With
    
C.Visible = True
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
            
Call Add_Win(Hwnd)
          
Temp_Str = Tempro_Read()
Tempro_Graph.Refresh
Call Send_Data_To_Graph(Tempro_Graph, 0, Temp_Str)

End Sub

Private Sub Form_Paint()
    Call First_Do
    If China_English = CHINA Then
        Exit_Command.Caption = "ESC退出"
         Print_Command.Caption = "&P图形输出"
        For axisID = VtChAxisIdX To VtChAxisIdZ
        With Tempro_Graph.Plot.Axis(axisID, 1).AxisTitle
            .Visible = True
            Select Case axisID
                Case 0
                    .Text = "时间 （秒)"
                Case 1
                    .Text = "温度 (度)"
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
        For axisID = VtChAxisIdX To VtChAxisIdZ
        With Tempro_Graph.Plot.Axis(axisID, 1).AxisTitle
            .Visible = True
            Select Case axisID
                Case 0
                    .Text = "Timer (S)"
                Case 1
                        .Text = "temrpr"
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
With Tempro_Graph
  .Left = 0: .Top = 0
  .Width = SSPanel2.Width
  .Height = SSPanel2.Height
End With
If China_English = CHINA Then
    Caption = "温度图"
Else
    Caption = "The Graphics of  Tempro"
End If
End Sub

Private Function Tempro_Read() As String
            Dim db      As Database
            Dim rs      As Recordset
            Dim sqlstr    As String
            
            Set db = OpenDatabase(Data_Path & "\make.mdb")
            
            sqlstr = "select    tempro   from grap_search"
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
            Tempro_Read = ""
            With rs
                Do While Not .EOF
                        If (Tempro_Read = "") Then
                                Tempro_Read = vFieldVal(.Fields("tempro").Value)
                        Else
                                Tempro_Read = Tempro_Read & "," & vFieldVal(.Fields("tempro").Value)
                        End If
                        .MoveNext
                Loop
         End With
         Debug.Print Tempro_Read
        rs.Close
        db.Close
End Function

Private Sub Print_Command_Click()
'Tempro_Graph.Print
End Sub


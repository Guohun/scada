VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form zhu 
   Caption         =   "Form1"
   ClientHeight    =   5436
   ClientLeft      =   660
   ClientTop       =   1020
   ClientWidth     =   8328
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
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
   ScaleHeight     =   5436
   ScaleWidth      =   8328
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   4
      Text            =   "1"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   3
      Text            =   "1"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox cb1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      ItemData        =   "zu_rmd.frx":0000
      Left            =   840
      List            =   "zu_rmd.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   3012
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   5040
      TabIndex        =   12
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
      Font3D          =   2
      Picture         =   "zu_rmd.frx":0004
   End
   Begin Threed.SSCommand Print_Report 
      Height          =   732
      Left            =   2040
      TabIndex        =   13
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
      Font3D          =   2
      MouseIcon       =   "zu_rmd.frx":0456
      Picture         =   "zu_rmd.frx":08A8
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "配方名"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   4680
      TabIndex        =   11
      Top             =   3480
      Width           =   1210
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "机器"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   4680
      TabIndex        =   10
      Top             =   2760
      Width           =   1210
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "班级"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   2040
      Width           =   1210
   End
   Begin VB.Label Label2 
      Caption         =   "打印日期选择"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   840
      TabIndex        =   8
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "End_day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "begin_day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1572
   End
End
Attribute VB_Name = "zhu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ban_hao As Integer
Public Begin_date, End_date As String
Public Search_Mathine As Integer
Public Search_Text As String
Public Search_Number As String




Dim StrBuffer As String * 250

Private Sub cb1_Click()
    If cb1.ListIndex = 0 Then                    'Trim(cb1.Text) = "指定日报表" Then
            Txt1(1).Visible = False
            Label4(1).Visible = False
            Label4(0).Caption = "日期"
            If Len(Begin_date) < 4 Then
                    Txt1(0).Text = Date$
            Else
                    Txt1(0).Text = Begin_date
            End If
    End If
    If cb1.ListIndex = 1 Then                                                  'Trim(cb1.Text) = "指定月报表" Then
             Label4(0).Caption = "月份："
             Label4(1).Caption = "年份："
    
            Txt1(1).Visible = True
            Label4(1).Visible = True
            Txt1(0).Text = Mid(Begin_date, 6, 2)
            If Len(End_date) < 4 Then
                    Txt1(1).Text = Year(Date)
            Else
                    Txt1(1).Text = Mid(End_date, 1, 4)
            End If
    End If
    If cb1.ListIndex = 2 Then                         'Trim(cb1.Text) = "指定某日到某日报表" Then
              Label4(0).Caption = "从某日开始："
             Label4(1).Caption = "到某日为终止"

            Txt1(1).Visible = True
            Label4(1).Visible = True
            If Len(Begin_date) < 4 Then
                    Txt1(0).Text = Date$
            Else
                    Txt1(0).Text = Begin_date
            End If
            If Len(End_date) < 6 Then
                    Txt1(1).Text = Date$
            Else
                    Txt1(1).Text = End_date
            End If
    End If
    
End Sub




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
    If IsEmpty(Begin_date) Then
                Txt1(0).Text = Date$
                Begin_date = Txt1(0).Text
    Else
                Txt1(0).Text = Begin_date
    End If
     If IsEmpty(End_date) Then
                Txt1(1).Text = Date$
                End_date = Txt1(1).Text
    Else
                Txt1(1).Text = End_date
    End If
    cb1.AddItem "this data of report"
    cb1.AddItem "the month of report"
    cb1.AddItem "from data to next data of report"
    cb1.Text = "this data of report"
    Call Add_Win(Hwnd)

End Sub

Private Sub Form_Paint()
Dim Temp_Idx As Integer
If China_English = CHINA Then
    Exit_Command.Caption = "ESC退出"
    Print_Report.Caption = "&P显示打印报表"
Else
    Exit_Command.Caption = "ESC"
    Print_Report.Caption = "&P Display Report "
End If
If China_English = CHINA Then
        If Me.Tag = 1 Then
          Me.Caption = "—— 产量统计报表——"
        Else
          Me.Caption = "—— 材料统计报表——"
        End If
          Label2(1).Caption = "统计方式选择："
          Label4(2).Caption = "班级"
          Label4(3).Caption = "机号"
          Label4(4).Caption = "配方号"

        

        
         Temp_Idx = cb1.ListIndex
         cb1.List(0) = "指定日报表"
         cb1.List(1) = "指定月报表"
         cb1.List(2) = "指定某日到某日报表"
         cb1.ListIndex = Temp_Idx

    Else
        If Me.Tag = 1 Then
          Me.Caption = "—— Produce total report——"
        Else
          Me.Caption = "—— Cai Liao  total report——"
        End If

        Label2(1).Caption = "the report name: "
        Label4(0).Caption = "from begin_day:"
        Label4(1).Caption = "print to this day:"

        Print_Report.Caption = "print report"
 End If


End Sub



Private Sub Form_Unload(Cancel As Integer)
Call Del_Win(Hwnd)

End Sub




Private Sub Print_Report_Click()
    Txt1(0).Enabled = True
    Txt1(1).Enabled = True
    If Trim(Txt1(0).Text) = "____-__-__" And Trim(Txt1(1).Text) = "____-__-__" Then
        MsgBox "Please input the begin day!", 48, "Prompt"
        Txt1(0).SetFocus
        Exit Sub
    End If
    
    If Trim(Txt1(0).Text) = "____-__-__" Then
        Txt1(0).Text = Txt1(0).Text
    End If
    If Trim(Txt1(0).Text) = "____-__-__" Then
        Txt1(0).Text = Txt1(0).Text
    End If
    
    Select Case cb1.ListIndex + 1
        Case 1
            If (Not IsDate(Txt1(0).Text)) Then
                    Speak_Error ("错误日期")
                    Exit Sub
            End If
            Begin_date = Txt1(0).Text
            Search_Text = "日报表"
        Case 2
            Begin_date = Val(Txt1(0).Text)
            If (Not IsNumeric(Txt1(0).Text)) Or Begin_date < 1 Or Begin_date > 12 Then
                    Speak_Error ("错误月份")
                    Exit Sub
            End If
            End_date = Val(Txt1(1).Text)
            If (Not IsNumeric(Txt1(1).Text)) Then
                    Speak_Error ("错误年份")
                    Exit Sub
            End If
            Search_Text = "月报表"
        Case 3
            If (Not IsDate(Txt1(0).Text)) Then
                    Speak_Error ("错误开始日期")
                    Exit Sub
            End If
            If (Not IsDate(Txt1(1).Text)) Then
                    Speak_Error ("错误结束日期")
                    Exit Sub
            End If
            Begin_date = Txt1(0).Text
            End_date = Txt1(1).Text
            Search_Text = "指定报表"
    End Select
        
        If Trim(Txt1(2).Text) <> "" Then
            If (Not IsNumeric(Txt1(2).Text)) Then
                    Speak_Error ("错误班号")
                    Txt1(2).SetFocus
                    Exit Sub
            End If
            Ban_hao = Txt1(2).Text
        Else
            Ban_hao = 0
        End If
        
        If Txt1(3).Text <> "" Then
            If (Not IsNumeric(Txt1(3).Text)) Then
                    Speak_Error ("错误机号")
                    Txt1(3).SetFocus
                    Exit Sub
            End If
                Search_Mathine = Txt1(3).Text
        Else
            Search_Mathine = 0
        End If
        Search_Number = Txt1(4).Text
        Print_Form.Show
End Sub

Private Sub Txt1_Change(index As Integer)
     If index = 0 Then Begin_date = Txt1(0).Text
End Sub

VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form zhu_pro 
   ClientHeight    =   5136
   ClientLeft      =   612
   ClientTop       =   1296
   ClientWidth     =   8316
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
   ScaleHeight     =   5136
   ScaleWidth      =   8316
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
      Left            =   6000
      TabIndex        =   5
      Text            =   "1"
      Top             =   3960
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
      TabIndex        =   0
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
      TabIndex        =   2
      Text            =   "1"
      Top             =   2640
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
      Top             =   3360
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
      Index           =   4
      Left            =   1800
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8310
      _Version        =   65536
      _ExtentX        =   14658
      _ExtentY        =   2143
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
         Height          =   735
         Left            =   5880
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   1110
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
         Picture         =   "zu_pro.frx":0000
      End
      Begin Threed.SSCommand Print_Report 
         Height          =   735
         Left            =   960
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   1110
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
         MouseIcon       =   "zu_pro.frx":0452
         Picture         =   "zu_pro.frx":08A4
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "车"
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
      Left            =   4680
      TabIndex        =   9
      Top             =   3960
      Width           =   1212
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
      TabIndex        =   8
      Top             =   2760
      Width           =   1572
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
      TabIndex        =   7
      Top             =   2640
      Width           =   1212
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
      TabIndex        =   6
      Top             =   3360
      Width           =   1212
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
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      Width           =   1212
   End
End
Attribute VB_Name = "zhu_pro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ban_hao As Integer
Public Now_Che As Integer
Public Begin_date   As String
Public Search_Mathine As Integer
Public Search_Text As String
Public Search_Number As String
Dim StrBuffer As String * 250





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
 
    Txt1(0).Text = Date$
    Call Add_Win(Hwnd)
End Sub

Private Sub Form_Paint()
  If China_English = CHINA Then
            Exit_Command.Caption = "ESC退出"
    Else
            Exit_Command.Caption = "Exit"
    End If

If China_English = CHINA Then
    If Me.Tag = "produce" Then
          Me.Caption = "—— 密炼过程报表——"
         Print_Report.Caption = "&P显示输出"
    End If
    If Me.Tag = "tempro" Then
            Me.Caption = "—— 温度对比图——"
         Print_Report.Caption = "&P图形输出"
    End If
    If Me.Tag = "power" Then
           Me.Caption = "—— 能量对比图——"
         Print_Report.Caption = "&P图形输出"
    End If
          Label4(0).Caption = "日期："
          Label4(1).Caption = "车号"
          Label4(2).Caption = "班级"
          Label4(3).Caption = "机号"
          Label4(4).Caption = "配方号"
    Else
        If Me.Tag = "produce" Then
            Me.Caption = "—— Produce统计报表——"
        Else
          

        End If

        Label4(0).Caption = "from begin_day:"
        Label4(1).Caption = "print to this day:"
        
        Print_Report.Caption = "print report"
        
    
 End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Del_Win(Hwnd)
End Sub





Private Sub Print_Report_Click()
Dim i%
    If Trim(Txt1(0).Text) = "____-__-__" Then
        MsgBox "Please input the begin day!", 48, "Prompt"
        Txt1(0).SetFocus
        Exit Sub
    End If
 For i% = 0 To 4
        If Txt1(i%) = "" Then
            Call Speak_Error("必须指定 " + Label4(i%).Caption)
            Exit Sub
        End If
 Next i%

      Now_Che = Txt1(1).Text
     Ban_hao = Txt1(2).Text
     Begin_date = Txt1(0).Text
    Search_Mathine = Txt1(3).Text
    Search_Number = Txt1(4).Text
  

        Select Case Me.Tag
        Case "power"
                Search_Text = "能量报表"
                 Power_Form.Show

        Case "tempro"
                Search_Text = "温度报表"
                 bl.Show
        Case "produce"
                Search_Text = "过程报表"
                 Produce_Form.Show
    End Select

    
    Exit Sub
End Sub



Private Sub Txt1_LostFocus(index As Integer)
                Select Case index
                    Case 1, 2, 3
                    If Val(Txt1(index)) = 0 Then
                        Speak_Error ("必须输入为大于0的数字 ")
                        Txt1(index).SetFocus
                        Exit Sub
                    End If
                    Txt1(index) = Val(Txt1(index))
                End Select
                

End Sub

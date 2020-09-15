VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Date_form 
   Caption         =   "date_Form"
   ClientHeight    =   4608
   ClientLeft      =   912
   ClientTop       =   1416
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4608
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Time_Text 
      Height          =   732
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   1692
   End
   Begin VB.TextBox Date_Text 
      Height          =   732
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1692
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   492
      Left            =   3000
      TabIndex        =   4
      Top             =   3720
      Width           =   852
      _Version        =   65536
      _ExtentX        =   1503
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "ESC"
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
      MousePointer    =   99
   End
   Begin Threed.SSCommand Ok_Command 
      Height          =   492
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "OK"
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
      MousePointer    =   99
   End
   Begin Threed.SSCommand Cancel_Command 
      Height          =   492
      Left            =   1800
      TabIndex        =   6
      Top             =   3720
      Width           =   972
      _Version        =   65536
      _ExtentX        =   1714
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "Change"
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
      MousePointer    =   99
   End
   Begin VB.Label Label2 
      Caption         =   "系统时间"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "修改系统日期:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1452
   End
End
Attribute VB_Name = "Date_form"
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
    Date_Text.Text = Date
    Time_Text.Text = time
    Call Add_Win(Hwnd)
End Sub

Private Sub Form_Paint()
If China_English = CHINA Then
     Caption = "系统时间"
    Ok_Command.Caption = "&O确认"
    Cancel_Command.Caption = "&C取消"
    Exit_Command.Caption = "ESC退出"
    Label1.Caption = "修改系统日期"
    Label2.Caption = "修改系统时间"
 Else
    Ok_Command.Caption = "&OK"
    Cancel_Command.Caption = "&Cancel"
    Exit_Command.Caption = "ESC Exit"
    Label1.Caption = "date"
    Label2.Caption = "time"
    
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
        Call Del_Win(Hwnd)
End Sub

Private Sub Ok_Command_Click()
On Error GoTo Exit_Do
  Date = Date_Text.Text
  time = Time_Text.Text
  Exit Sub
Exit_Do:
    Call Speak_Error("错误")
    Date_Text.SetFocus
End Sub

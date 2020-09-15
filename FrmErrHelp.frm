VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmErrHelp 
   Caption         =   "Form1"
   ClientHeight    =   4512
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6168
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4512
   ScaleWidth      =   6168
   Begin VB.TextBox ErrorHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3012
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "FrmErrHelp.frx":0000
      Top             =   1200
      Width           =   5652
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
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
      Font3D          =   1
      Picture         =   "FrmErrHelp.frx":0006
   End
   Begin Threed.SSCommand Update_Command 
      Height          =   732
      Left            =   1680
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "NoEdit"
      Top             =   240
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
      Font3D          =   1
      Picture         =   "FrmErrHelp.frx":0458
   End
End
Attribute VB_Name = "frmErrHelp"
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
            Width = MAIN_MDI.Width / 2 + 2000
' Set width of form.
            Height = MAIN_MDI.Height / 2 + 2000
 ' Set height of form.
            Left = 1000
            Top = 1000
            Call Add_Win(Hwnd)

End Sub

Private Sub Form_Paint()
        If China_English = CHINA Then
                Caption = "´íÎó¹ÊÕÏ°ïÖú"
                Exit_Command.Caption = "ESCÍË³ö"
                Update_Command.Caption = "F5È·¶¨"
        Else
                Caption = "help"
                Exit_Command.Caption = "ESC "
                Update_Command.Caption = "F5 "
        
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Del_Win(Hwnd)
End Sub

Private Sub Update_Command_Click()
    If frmErrHelp.Tag = "china_help" Then
        error_Input!DError_num_table.Recordset.Edit
         error_Input!DError_num_table.Recordset.Fields("china_help").Value = ErrorHelp.Text
         error_Input!DError_num_table.Recordset.Update
    Else
        error_Input!DError_num_table.Recordset.Edit
        error_Input!DError_num_table.Recordset.Fields("english_help").Value = ErrorHelp.Text
        error_Input!DError_num_table.Recordset.Update
    End If

End Sub

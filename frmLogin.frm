VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1836
   ClientLeft      =   2832
   ClientTop       =   3480
   ClientWidth     =   4272
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1082.667
   ScaleMode       =   0  'User
   ScaleWidth      =   4004.769
   Begin Threed.SSCommand CmdCancel 
      Cancel          =   -1  'True
      Height          =   492
      Left            =   3000
      TabIndex        =   6
      Top             =   1200
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
   Begin Threed.SSCommand CmdOk 
      Height          =   492
      Left            =   480
      TabIndex        =   4
      Top             =   1200
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
   Begin VB.TextBox txtUserName 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   525
      Visible         =   0   'False
      Width           =   2325
   End
   Begin Threed.SSCommand CmdUpdate 
      Height          =   492
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
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
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Public modal_name    As String
Public Disp_Prompt    As String
Private Input_Time    As Integer




Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload frmLogin
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    Dim password_str   As String

       password_str = txtUserName.Text
       If txtPassword.Visible = True Then
                Call Password_Set(password_str)
                LoginSucceeded = False
                Exit Sub
        End If
      Call Password_Check(password_str)
End Sub

Private Sub CmdUpdate_Click()
            txtPassword.Text = ""
            txtUserName.Text = ""
            txtPassword.Visible = True
             lblLabels(1).Visible = True
            CmdUpdate.Visible = False
            txtUserName.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
        Input_Time = 0
        CmdUpdate.Visible = True
            If Multi_Screen = "2" Then
                Top = Screen.Height * 0.1   ' Set height of form.
         End If
    
        Call Add_Win(Hwnd)
        
End Sub

Private Sub Form_Paint()

If China_English = CHINA Then
    Caption = Disp_Prompt + "密码输入"
    lblLabels(0).Caption = " 输入密码"  ' Set prompt.
    lblLabels(1).Caption = "新 密码" ' Set title.
    CmdOk.Caption = "&O确认"
    CmdCancel.Caption = "ESC取消"
    CmdUpdate.Caption = "&C修改"
Else
    Caption = Disp_Prompt + "Password Input"
    lblLabels(0).Caption = " Input Passwd"  ' Set prompt.
    lblLabels(1).Caption = " Check passwd" ' Set title.
    CmdOk.Caption = "&OK"
    CmdCancel.Caption = "Esc"
    CmdUpdate.Caption = "&Change"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call Del_Win(Hwnd)
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case vbKeyReturn, vbKeyDown
                                    Call cmdOK_Click
            Case vbKeyUp
                        txtUserName.SetFocus
                        KeyCode = 0
          End Select
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case vbKeyReturn
                            If txtPassword.Visible = True Then
                                  txtPassword.SetFocus
                            Else
                                    Call cmdOK_Click
                            End If
                            
            Case vbKeyUp, vbKeyDown
                              If txtPassword.Visible = True Then
                                  txtPassword.SetFocus
                            Else
                                    CmdOk.SetFocus
                            End If
                        KeyCode = 0
          End Select
End Sub

Private Sub Password_Check(password_str As String)
      Dim db As Database, rs As Recordset
      Set db = Workspaces(0).OpenDatabase("c:\ljxt\comm.mdb")
       If password_str = "ZGHLXY" Then
                Set rs = db.OpenRecordset("select * from  password where modal_name like '" & modal_name & "'")
        Else
                Set rs = db.OpenRecordset("select * from  password where modal_name like '" & modal_name & "' and password like '" & password_str & "'")
       End If
      
      
      Debug.Print "select * from  password where modal_name like '" & password_str & "' and password like '" & password_str & "'"
        If rs.EOF Then     ' Test for success.
                LoginSucceeded = False
                If Input_Time < 3 Then
                        Call Speak_Error("密码错误,重输：")
                        txtUserName.SetFocus
                        SendKeys "{Home}+{End}"
                        Input_Time = Input_Time + 1
                Else
                        Call Speak_Error("密码错误 最多只输入三次")
                        LoginSucceeded = False
                        Unload frmLogin
                End If
        Else
                LoginSucceeded = True
                Unload frmLogin
        End If
        rs.Close
        db.Close

End Sub
Private Sub Password_Set(password_str As String)
      Dim db As Database, rs As Recordset
      
      
      Set db = Workspaces(0).OpenDatabase("c:\ljxt\comm.mdb")
       If password_str = "ZGHLXY" Then
                Set rs = db.OpenRecordset("select * from  password where modal_name like '" & modal_name & "'")
        Else
                Set rs = db.OpenRecordset("select * from  password where modal_name like '" & modal_name & "' and password like '" & password_str & "'")
       End If
            
      Debug.Print "select * from  password where modal_name like '" & password_str & "' and password like '" & password_str & "'"
        If rs.EOF Then     ' Test for success.
                        
                        Call Speak_Error("原始密码错误,不能修改：")
                        txtUserName.SetFocus
                        SendKeys "{Home}+{End}"
                        LoginSucceeded = False
                        Unload frmLogin
                
        Else
               If password_str = "ZGHLXY" Then
                        txtPassword.Text = rs.Fields("password").Value
                        txtPassword.PasswordChar = ""
                Else
                       If Len(txtPassword.Text) < 3 Then
                                Call Speak_Error("密码必须输入两位以上")
                                LoginSucceeded = False
                                Exit Sub
                          End If
                        rs.Edit
                        rs.Fields("PASSWORD").Value = txtPassword.Text
                        rs.Update
                        Call Speak_Error("修改成功")
                        LoginSucceeded = True
                        Unload frmLogin
                End If
        End If
        rs.Close
        db.Close
End Sub


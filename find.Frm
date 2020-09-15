VERSION 5.00
Begin VB.Form frmFind 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Record"
   ClientHeight    =   2436
   ClientLeft      =   2376
   ClientTop       =   1776
   ClientWidth     =   5448
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016128
   Icon            =   "FIND.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2411.999
   ScaleMode       =   0  'User
   ScaleWidth      =   5476.842
   Begin VB.ListBox lstFields 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00008000&
      Height          =   1248
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1692
   End
   Begin VB.OptionButton optFindType 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "FindLast"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton optFindType 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "FindPrevious"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton optFindType 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "FindNext"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.OptionButton optFindType 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "FindFirst"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.ListBox lstOperators 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00008000&
      Height          =   1248
      ItemData        =   "FIND.frx":030A
      Left            =   2040
      List            =   "FIND.frx":0323
      TabIndex        =   2
      Top             =   360
      Width           =   960
   End
   Begin VB.TextBox txtExpression 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   372
      Left            =   600
      TabIndex        =   8
      Top             =   1919
      Width           =   1691
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   2879
      TabIndex        =   9
      Top             =   1919
      Width           =   1691
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "Operators: "
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "Fields: "
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "Value or Expression: "
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbNotFound As Integer

   
Private Sub cmdCancel_Click()
    gbFindFailed = True
    Unload Me
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Paint()
  If China_English = CHINA Then
    Me.Caption = "记录查找"
    lblLabels(1).Caption = "关键字"
    lblLabels(0).Caption = "操作符"
    lblLabels(2).Caption = "查找值"
    optFindType(0).Caption = "查找第一个"
    optFindType(1).Caption = "查找下一个"
    optFindType(2).Caption = "查找前一个"
    optFindType(3).Caption = "查找最后一个"
    CmdOk.Caption = "开始查找"
    cmdCancel.Caption = "取消查找"
Else
    Me.Caption = "Search record"
    lblLabels(1).Caption = "Key word"
    lblLabels(0).Caption = "Option"
    lblLabels(2).Caption = "Search"
    optFindType(0).Caption = "FindFirst"
    optFindType(1).Caption = "FindNext"
    optFindType(2).Caption = "FindPrev"
    optFindType(3).Caption = "FindLast"
    CmdOk.Caption = "Begin Serch"
    cmdCancel.Caption = "Cancel"
  End If

End Sub

Private Sub Form_Terminate()
Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call Del_Win(Hwnd)
End Sub

Private Sub lstFields_GotFocus()
        If lstFields.ListCount > 0 And lstFields.ListIndex < 0 Then
                lstFields.ListIndex = 0
        End If
End Sub

Private Sub lstFields_KeyDown(KeyCode As Integer, Shift As Integer)
       Select Case KeyCode
            Case vbKeyReturn
                                  lstOperators.SetFocus
                                  KeyCode = 0
                            
            Case vbKeyLeft, vbKeyRight
                        lstOperators.SetFocus
                        KeyCode = 0
          End Select
End Sub

Private Sub lstOperators_GotFocus()
        If lstOperators.ListCount > 0 And lstOperators.ListIndex < 0 Then
                lstOperators.ListIndex = 0
        End If
End Sub

Private Sub lstOperators_KeyDown(KeyCode As Integer, Shift As Integer)
       Select Case KeyCode
            Case vbKeyReturn
                                  txtExpression.SetFocus
                                  KeyCode = 0
                            
            Case vbKeyLeft, vbKeyRight
                        txtExpression.SetFocus
                        KeyCode = 0
          End Select

End Sub

Private Sub optFindType_Click(index As Integer)
     '       gnFindType = index
End Sub

Private Sub txtExpression_Change()
  If Len(lstFields.Text) > 0 And Len(lstOperators.Text) > 0 And Len(txtExpression.Text) > 0 Then
    CmdOk.Enabled = True
  Else
    CmdOk.Enabled = False
  End If
End Sub

Private Sub lstFields_Click()
  If Len(lstFields.Text) > 0 And Len(lstOperators.Text) > 0 And Len(txtExpression.Text) > 0 Then
    CmdOk.Enabled = True
  Else
    CmdOk.Enabled = False
  End If
End Sub

Private Sub lstOperators_Click()
  If Len(lstFields.Text) > 0 And Len(lstOperators.Text) > 0 And Len(txtExpression.Text) > 0 Then
    CmdOk.Enabled = True
  Else
    CmdOk.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  'CenterMe Me, gnMDIFORM
  Call Add_Win(Hwnd)

  mbNotFound = False
End Sub

Private Sub cmdOK_Click()
   On Error GoTo FindErr

   Dim i As Integer

   i = lstFields.ListIndex
   mbNotFound = False
   'SetHourglass

   gsFindFiel = lstFields.Text
   gsFindExpr = txtExpression.Text
   gsFindOp = lstOperators.Text
   lstI = lstFields.ListIndex
   
   Screen.MousePointer = vbDefault
   Unload Me
   'Tan_hei_table_Input.Show
   Exit Sub

FindErr:
   'If Err <> gnEOF_ERR Then
   '  ShowError
     Exit Sub
   'Else
   '  mbNotFound = True
   '  Resume Next
   'End If

End Sub

Private Sub txtExpression_KeyDown(KeyCode As Integer, Shift As Integer)
       Select Case KeyCode
            Case vbKeyReturn
                                    Call cmdOK_Click
                                  KeyCode = 0
          End Select
End Sub

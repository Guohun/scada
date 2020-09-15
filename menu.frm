VERSION 5.00
Begin VB.Form Menu_table_Input 
   Caption         =   "5"
   ClientHeight    =   5676
   ClientLeft      =   456
   ClientTop       =   1080
   ClientWidth     =   8316
   KeyPreview      =   -1  'True
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5676
   ScaleWidth      =   8316
   Begin VB.Frame Frame2 
      Height          =   732
      Left            =   0
      TabIndex        =   19
      Top             =   4200
      Width           =   7692
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&F<<"
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&+"
         Height          =   375
         Left            =   5280
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&-"
         Height          =   375
         Left            =   6000
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&L>>"
         Height          =   375
         Left            =   6720
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.Label record_prompt 
         Caption         =   "Label1"
         Height          =   372
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1812
      End
   End
   Begin VB.CommandButton Exit_Command 
      BackColor       =   &H000000C0&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H000000C0&
      Caption         =   "&Find Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000C0&
      Caption         =   "&Del Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H000000C0&
      Caption         =   "&Add  Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   7692
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "key"
         DataSource      =   "DMenu_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   2
         Top             =   360
         Width           =   1655
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "english_name"
         DataSource      =   "DMenu_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   5
         Left            =   5760
         TabIndex        =   6
         Top             =   2640
         Width           =   1655
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "china_name"
         DataSource      =   "DMenu_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   5
         Top             =   2640
         Width           =   1655
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "col"
         DataSource      =   "DMenu_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   3
         Left            =   5760
         TabIndex        =   4
         Top             =   1560
         Width           =   1655
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "row"
         DataSource      =   "DMenu_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   1655
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "father_name"
         DataSource      =   "DMenu_table"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1655
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Key:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   372
         Index           =   1
         Left            =   4440
         TabIndex        =   17
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "English_name:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   372
         Index           =   5
         Left            =   4440
         TabIndex        =   16
         Top             =   2640
         Width           =   1332
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "China_name:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   372
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   1332
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Col:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   372
         Index           =   3
         Left            =   4440
         TabIndex        =   12
         Top             =   1560
         Width           =   1332
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Row:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   372
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1332
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Father_name:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   372
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1452
      End
   End
   Begin VB.Data DMenu_table 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\ljxt.Mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "menu_table"
      Top             =   3360
      Visible         =   0   'False
      Width           =   6735
   End
End
Attribute VB_Name = "Menu_table_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msTBookMark As String
Dim gsDBName As Database
Dim gsThisSet As Recordset
Private Sub cmdAddNew_Click()
 On Error GoTo AddErr
 ChangeButtons
 DMenu_table.Recordset.AddNew
 If DMenu_table.Recordset.RecordCount > 0 Then
   msTBookMark = DMenu_table.Recordset.Bookmark
 Else
   msTBookMark = ""
  End If
AddErr:
    Exit Sub
End Sub


Private Sub cmdCancel_Click()
DMenu_table.Recordset.CancelUpdate
RestoreButtons
End Sub


Private Sub cmdDelete_Click()
If China_English = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
  DMenu_table.Recordset.Delete
  Call cmdPrev_Click
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then DMenu_table.Recordset.Delete
End If
End Sub

Private Sub cmdEdit_Click()
DMenu_table.Recordset.Edit
ChangeButtons
End Sub

Private Sub CmdFind_Click()
On Error GoTo FindErr:
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
FindStart:
  For i = 0 To 5
      frmFind.lstFields.AddItem Label(i).Caption 'Mid((Label(i).Caption), 1, Len(Label(i).Caption) - 1)
  Next
  gbFindFailed = False
  gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOp
  frmFind.txtExpression.Text = gsFindExpr
  frmFind.optFindType(gnFindType) = True

  frmFind.Show vbModal
 If gbFindFailed = True Then Exit Sub 'find cancelled
  Select Case gsFindFiel
    Case Label(0).Caption
      sTmp = txtFields(0).DataField + space(1) + gsFindOp + space(1) + "'" + gsFindExpr + "'"
    Case Label(1).Caption
      sTmp = txtFields(1).DataField + space(1) + gsFindOp + space(1) + "'" + gsFindExpr + "'"
    Case Label(2).Caption
      sTmp = txtFields(2).DataField + space(1) + gsFindOp + space(1) + gsFindExpr
    Case Label(3).Caption
      sTmp = txtFields(3).DataField + space(1) + gsFindOp + space(1) + gsFindExpr
    Case Label(4).Caption
      sTmp = txtFields(5).DataField + space(1) + gsFindOp + space(1) + "'" + gsFindExpr + "'"
    Case Label(5).Caption
      sTmp = txtFields(6).DataField + space(1) + gsFindOp + space(1) + "'" + gsFindExpr + "'"
      
  End Select
  i = frmFind.lstFields.ListIndex
  sBookMark = DMenu_table.Recordset.Bookmark
  'search for the record
  Select Case gnFindType
    Case 0
      DMenu_table.Recordset.FindFirst sTmp
    Case 1
      DMenu_table.Recordset.FindNext sTmp
    Case 2
      DMenu_table.Recordset.FindPrevious sTmp
    Case 3
      DMenu_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = DMenu_table.Recordset.NoMatch
  If mbNotFound = True Then Speak_Error ("not found!")
AfterWhile:

  Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    DMenu_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    DMenu_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = DMenu_table.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    DMenu_table.Recordset.Bookmark = sBookMark
  End If

  mbDataChanged = False
  txtFields(0).SetFocus
  Exit Sub

FindErr:
'  Screen.MousePointer = vbDefault
'  If Err <> gnEOF_ERR Then
'    ShowError
  txtFields(0).SetFocus
    Exit Sub
''  Else
    mbNotFound = True
'    Resume Next
'  End If

End Sub




Private Sub cmdFirst_Click()
  DMenu_table.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
DMenu_table.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
DMenu_table.Recordset.MoveNext
If DMenu_table.Recordset.EOF Then DMenu_table.Recordset.MoveLast
End Sub

Private Sub cmdPrev_Click()
DMenu_table.Recordset.MovePrevious
If DMenu_table.Recordset.BOF Then DMenu_table.Recordset.MoveFirst
End Sub


Private Sub CmdUpdate_Click()
'On Error GoTo UpdateErr

'retryupd:
 Dim i%
    For i% = 0 To 5
        If txtFields(i%) = "" Then
                        Call Speak_Error(Label(i%).Caption & "不能为空")
                        txtFields(i%).SetFocus
                        Exit Sub
        End If
  Next i%
  DMenu_table.Recordset.Update
  RestoreButtons
  DMenu_table.Recordset.Bookmark = msTBookMark
  Exit Sub

'UpdateErr:
  'check for locked error
'  If Err = 3260 And nRetryCnt < gnMURetryCnt Then
'    nRetryCnt = nRetryCnt + 1
'    DMenu_table.Recordset.Bookmark = msBookMark   'Cancel the update
'    DBEngine.Idle dbFreeLocks
'    nDelay = Timer
    'Wait gnMUDelay seconds
'    While Timer - nDelay < gnMUDelay
 '     'do nothing
'    Wend
'    Resume retryupd
'  Else
    'ShowError
 '   Exit Sub
 ' End If

End Sub




Public Sub ChangeButtons()
  cmdAddNew.Visible = False
'  cmdEdit.visible = False
'  cmdFind.Visible = False
  cmdDelete.Visible = False
  CmdUpdate.Visible = True
  CmdCancel.Visible = True
  'combo1.Top = txtFields(2).Top
  'combo1.Left = txtFields(2).Left
  'txtFields(2).Visible = False
  'Label13.Top = Frame2.Top
  'Frame2.Visible = False
  cmdFirst.Visible = False
  cmdPrev.Visible = False
  cmdNext.Visible = False
  cmdLast.Visible = False
  'Label14.Top = Frame2.Top
  'Label15.Top = Frame2.Top
  'Combo2.Top = Frame2.Top + 260
  'combo1.Visible = True
  'Label13.Visible = True
  'Label14.Visible = True
  'Label15.Visible = True
  'Combo2.Visible = True
End Sub

Public Sub RestoreButtons()
  cmdAddNew.Visible = True
  cmdDelete.Visible = True
'  cmdFind.Visible = True
  CmdUpdate.Visible = False
  CmdCancel.Visible = False
  'txtFields(2).Visible = True
  'combo1.Visible = False
  'combo1.Visible = False
  'Label13.Visible = False
  'Label14.Visible = False
  'Label15.Visible = False
  'Combo2.Visible = False
  Frame2.Visible = True
  cmdFirst.Visible = True
  cmdPrev.Visible = True
  cmdNext.Visible = True
  cmdLast.Visible = True

End Sub
Private Sub DMenu_table_Reposition()
If China_English = CHINA Then
  If Not DMenu_table.Recordset.EOF And Not DMenu_table.Recordset.BOF Then
    record_prompt.Caption = "当前记录:" & (DMenu_table.Recordset.AbsolutePosition + 1) & " 之 " & DMenu_table.Recordset.RecordCount
  Else
      record_prompt.Caption = "新记录:"
  End If
Else

End If

End Sub

Private Sub DMenu_table_Validate(Action As Integer, Save As Integer)
   If Save = -1 Then
    For i% = 0 To 5
        If txtFields(i%) = "" Then
                        Call Speak_Error(Label(i%).Caption & "不能为空")
                         Save = 0
                         Action = 0
                        txtFields(i%).SetFocus
                        Exit Sub
        End If
  Next i%
End If
End Sub


Private Sub Exit_Command_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    Dim db As Database, rs As Recordset
      Set db = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.mdb")
      Set rs = db.OpenRecordset("select * from menu_table order by  father_name ,col")
      Set DMenu_table.Recordset = rs
            Width = MAIN_MDI.Width  ' Set width of form.
            Height = MAIN_MDI.Height * 3 / 4 ' Set height of form.
            Left = 0
            Top = 0
            Call Add_Win(Hwnd)

End Sub

Public Sub FirstDo()
Dim i, count As Integer

End Sub

Private Sub Form_Paint()
If China_English = True Then
   Caption = "菜单管理"
  cmdAddNew.Caption = "&A添加记录"
'  cmdEdit.Caption = "编辑记录(&E)"
  cmdDelete.Caption = "&D删除记录"
  CmdFind.Caption = "&F查询记录"
  CmdUpdate.Caption = "&U更新"
  CmdCancel.Caption = "&C取消"
  Exit_Command.Caption = "&E退出"
  Frame1.Caption = ""
  Frame2.Caption = "移动记录"
'  Label(0).Caption = "菜单标志:"
  Label(0).Caption = "父菜单标志:"
'  Label(2).Caption = "后菜单标志:"
 ' Label(3).Caption = "前菜单标志:"
   Label(1).Caption = "快捷键:"
  Label(2).Caption = "行:"
  Label(3).Caption = "列:"
  Label(4).Caption = "中文名:"
  Label(5).Caption = "英文名:"
  
'  Label(9).Caption = "标志:"
 ' Label(10).Caption = "菜单:"
  'Label(11).Caption = "备用:"
  Else
    Caption = "Menu Manage"
End If

End Sub

Private Sub Form_Terminate()
Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call Del_Win(Hwnd)
End Sub


Private Sub txtFields_GotFocus(index As Integer)
        txtFields(index).ForeColor = Get_Focus_Fore_Color
End Sub


Private Sub txtFields_LostFocus(index As Integer)
        txtFields(index).ForeColor = Los_Focus_Fore_Color

End Sub



VERSION 4.00
Begin VB.Form Lian_jiao_table_Input 
   Caption         =   "Lian_liao_table_Input:"
   ClientHeight    =   4368
   ClientLeft      =   1428
   ClientTop       =   1428
   ClientWidth     =   6768
   BeginProperty Font 
      name            =   "System"
      charset         =   0
      weight          =   700
      size            =   9.6
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   4752
   Icon            =   "lian_jiao_table.frx":0000
   Left            =   1380
   LinkTopic       =   " "
   MaxButton       =   0   'False
   ScaleHeight     =   4368
   ScaleWidth      =   6768
   Top             =   1092
   Width           =   6864
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0FFC0&
      Height          =   360
      Left            =   3480
      TabIndex        =   35
      Text            =   "Combo2"
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1320
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H000000C0&
      Caption         =   "&Find Record"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000C0&
      Caption         =   "&Del Record"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H000000C0&
      Caption         =   "&Edit Record"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H000000C0&
      Caption         =   "&Add  Record"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Frame1"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   7.8
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6735
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "ctr"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   10
         Left            =   4920
         TabIndex        =   15
         Text            =   "txtCtr"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "neg"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   9
         Left            =   1560
         TabIndex        =   14
         Text            =   "txtNeg"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFC0&
         Height          =   360
         Left            =   4920
         TabIndex        =   34
         Text            =   "Combo1"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Data BTable 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\ljxt\Ljxt.mdb"
         Exclusive       =   0   'False
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   7.8
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "cai_liao_table"
         Top             =   3360
         Width           =   1140
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "number"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Text            =   "txtNumber"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "tem"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   8
         Left            =   4920
         TabIndex        =   13
         Text            =   "txtTem"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "time"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   7
         Left            =   1560
         TabIndex        =   12
         Text            =   "txtTime"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Data Dlian_jiao_table 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "c:\ljxt\ljxt.Mdb"
         Exclusive       =   0   'False
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   0
            weight          =   400
            size            =   7.8
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tan_hei_table"
         Top             =   3360
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "pre"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   6
         Left            =   4920
         TabIndex        =   11
         Text            =   "txtPre"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "speed"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   5
         Left            =   1560
         TabIndex        =   10
         Text            =   "txtSpeed"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "st"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   4
         Left            =   4920
         TabIndex        =   9
         Text            =   "txtSt"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "tou"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   8
         Text            =   "txtTou"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "du2"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   7
         Text            =   "txtName"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFC0&
         DataField       =   "du1"
         DataSource      =   "Dlian_jiao_table"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Text            =   "txtDu1"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ctr:"
         Height          =   375
         Index           =   10
         Left            =   3600
         TabIndex        =   37
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Neg:"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   36
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "number:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Tem:"
         Height          =   375
         Index           =   8
         Left            =   3600
         TabIndex        =   32
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Time:"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   31
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Pre:"
         Height          =   375
         Index           =   6
         Left            =   3600
         TabIndex        =   25
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Speed:"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   24
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "At:"
         Height          =   375
         Index           =   4
         Left            =   3600
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Tou:"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Du2:"
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   21
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H0000C0C0&
         Caption         =   "Du1:"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Move Record"
      Height          =   735
      Left            =   3840
      TabIndex        =   29
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Rows"
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Position of Record:"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   3135
   End
End
Attribute VB_Name = "lian_jiao_table_Input"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim msTBookMark As String
'Dim china_english As Integer
Dim gsDBName As Database
Dim gsThisSet As Recordset
Dim gsThatSet As Recordset
Function StripNonAscii(rvntVal As Variant) As String
  Dim i As Integer
  Dim sTmp As String

  For i = 1 To Len(rvntVal)
    If Asc(Mid(rvntVal, i, 1)) < 32 Or Asc(Mid(rvntVal, i, 1)) > 126 Then
      sTmp = sTmp & " "
    Else
      sTmp = sTmp & Mid(rvntVal, i, 1)
    End If
  Next

  StripNonAscii = sTmp

End Function




Function vFieldVal(rvntFieldVal As Variant) As Variant
  If IsNull(rvntFieldVal) Then
    vFieldVal = ""
  Else
    
    vFieldVal = CStr(rvntFieldVal)
  End If
End Function

Private Sub cmdAddNew_Click()
 On Error GoTo Adderr
 ChangeButtons
   'DBCombo1.ListField = "number"

 'set the mode
 'FAddNewFlag = True
 'cmdFirst.Enabled = False
 'cmdPrev.Enabled = False
 'cmdNext.Enabled = False
 'cmdlast.Enabled = False
 Dlian_jiao_table.Recordset.AddNew
 'For i = 0 To 8 'Dlian_jiao_table.Recordset.Fields.Count - 1
 '   txtFields(i).Text = ""
 'Next
 combo1.ListIndex = 0
 Call Combo1_Click
 'txtFields(1).Text = gsThisSet.Fields("number").Value
 'txtFields(3).Text = gsThisSet.Fields("dou").Value
 Dlian_jiao_table.Recordset.Fields("user_name").Value = user_name
 Dlian_jiao_table.Recordset.Fields("Password").Value = password
 ' mbAddNewFlag = True
 If Dlian_jiao_table.Recordset.RecordCount > 0 Then
   msTBookMark = Dlian_jiao_table.Recordset.Bookmark
 Else
   msTBookMark = ""
  End If
Adderr:
    Exit Sub
End Sub


Private Sub cmdCancel_Click()
Dlian_jiao_table.Recordset.CancelUpdate
'Dlian_jiao_table.Recordset.Bookmark = msTBookMark
RestoreButtons
End Sub


Private Sub cmdDelete_Click()
If china_english = True Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
  Dlian_jiao_table.Recordset.Delete
  Call cmdPrev_Click
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then Dlian_jiao_table.Recordset.Delete
End If
End Sub

Private Sub cmdEdit_Click()
Dim LIndex As Integer
Dlian_jiao_table.Recordset.Edit
ChangeButtons
LIndex = 0
If txtFields(2).Text <> "" Then
Do
 If (txtFields(2).Text) = combo1.List(LIndex) Then Exit Do
 LIndex = LIndex + 1
Loop
End If
combo1.ListIndex = LIndex
Call Combo1_Click
End Sub

Private Sub cmdFind_Click()
 On Error GoTo FindErr
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String
  'load the column names into the find form
FindStart:
For i = 0 To 8
      frmFind.lstFields.AddItem Mid((Label(i).Caption), 1, Len(Label(i).Caption) - 1)
  Next
    'reset the flags
  gbFindFaile = False
  gbFromTableView = False
  frmFind.lstFields.Text = gsFindFiel
  frmFind.lstOperators.Text = gsFindOpen
  frmFind.txtExpression.Text = gsFindExpror
  frmFind.optFindType(gsFindType) = True
'  mbNotFound = False
  frmFind.Show
  'If gbFindFailed = True Then   'find cancelled
  '  GoTo AfterWhile
  'End If
  If True Then Exit Sub
  Select Case gsFindFiel
    Case number Or 配方编号:
      sTmp = "number" + gsFindOpen + gsFindExpror
    Case cai_number Or 材料编号:
    sTmp = "number" + gsFindOpen + "'" + gsFindExpror + "'"
    Case name Or 材料名称:
    sTmp = "name" + gsFindOpen + "'" + gsFindExpror + "'"
    Case dou Or 斗号:
    sTmp = "dou" + gsFindOpen + "'" + gsFindExpror + "'"
    Case weight Or 所需重量:
    sTmp = "weight" + gsFindOpen + gsFindExpror
    Case fast_do Or 快慢提前量:
    sTmp = "fast_do" + gsFindOpen + gsFindExpror
    Case drop_do Or 点动提前量:
    sTmp = "drop_do" + gsFindExpror
    Case gon_cha Or 允许公差:
    sTmp = "gon_cha" + gsFindOpen + gsFindExpror
    Case time Or 控制时间:
    sTmp = "time" + gsFindOpen + gsFindExpror
  End Select
  i = frmFind.lstFields.ListIndex
  sBookMark = Dlian_jiao_table.Recordset.Bookmark
  'search for the record
  Select Case gsFindType
    Case 0
      Dlian_jiao_table.Recordset.FindFirst sTmp
    Case 1
      Dlian_jiao_table.Recordset.FindNext sTmp
    Case 2
      Dlian_jiao_table.Recordset.FindPrevious sTmp
    Case 3
      Dlian_jiao_table.Recordset.FindLast sTmp
  End Select
  mbNotFound = Dlian_jiao_table.Recordset.NoMatch

AfterWhile:

  Screen.MousePointer = vbDefault

  If gbFindFaile = True Then   'go back to original row
    Dlian_jiao_table.Recordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    Dlian_jiao_table.Recordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = Dlian_jiao_table.Recordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
    Dlian_jiao_table.Recordset.Bookmark = sBookMark
  End If

  mbDataChanged = False

  Exit Sub

FindErr:
'  Screen.MousePointer = vbDefault
'  If Err <> gnEOF_ERR Then
'    ShowError
    Exit Sub
''  Else
    mbNotFound = True
'    Resume Next
'  End If

End Sub




Private Sub cmdFirst_Click()
  Dlian_jiao_table.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
Dlian_jiao_table.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
Dlian_jiao_table.Recordset.MoveNext
If Dlian_jiao_table.Recordset.EOF Then Dlian_jiao_table.Recordset.MoveLast
End Sub

Private Sub cmdPrev_Click()
Dlian_jiao_table.Recordset.MovePrevious
If Dlian_jiao_table.Recordset.BOF Then Dlian_jiao_table.Recordset.MoveFirst
End Sub


Private Sub cmdUpdate_Click()
'On Error GoTo UpdateErr

'retryupd:
  Dlian_jiao_table.Recordset.Update
  RestoreButtons
  Dlian_jiao_table.Recordset.Bookmark = msTBookMark
  Exit Sub

'UpdateErr:
  'check for locked error
'  If Err = 3260 And nRetryCnt < gnMURetryCnt Then
'    nRetryCnt = nRetryCnt + 1
'    Dlian_jiao_table.Recordset.Bookmark = msBookMark   'Cancel the update
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
  cmdEdit.Visible = False
  cmdFind.Visible = False
  cmdDelete.Visible = False
  cmdUpdate.Visible = True
  cmdCancel.Visible = True
  'Combo2.Top = txtFields(1).Top
  'Combo2.Left = txtFields(1).Left
  'txtFields(1).Visible = False
  'Combo2.Visible = True
  'Combo1.Top = txtFields(2).Top
  'Combo1.Left = txtFields(2).Left
  'txtFields(2).Visible = False
  Frame2.Visible = False
  Me.Height = Me.Height - 740
  cmdFirst.Visible = False
  cmdPrev.Visible = False
  cmdNext.Visible = False
  cmdLast.Visible = False
  'Combo1.Visible = True
End Sub

Public Sub RestoreButtons()
  cmdAddNew.Visible = True
  cmdEdit.Visible = True
  cmdDelete.Visible = True
  cmdFind.Visible = True
  cmdUpdate.Visible = False
  cmdCancel.Visible = False
  'txtFields(1).Visible = False
 ' txtFields(2).Visible = True
 ' Combo1.Visible = False
 ' Combo2.Visible = False
 ' Me.Height = Me.Height + 740
  Frame2.Visible = True
  cmdFirst.Visible = True
  cmdPrev.Visible = True
  cmdNext.Visible = True
  cmdLast.Visible = True

End Sub




Private Sub Combo1_Click()
gsThisSet.FindFirst "name='" + combo1.List(combo1.ListIndex) + "'"
txtFields(1).Text = gsThisSet.Fields("number").Value
txtFields(3).Text = gsThisSet.Fields("dou").Value
Combo2.ListIndex = combo1.ListIndex
txtFields(2).Text = combo1.List(combo1.ListIndex)
End Sub

Private Sub DBCombo1_Change()

End Sub

Private Sub Combo2_Click()
'Dim i, RecordCount As Integer
'gsThisSet.MoveLast
'RecordCount = gsThisSet.RecordCount
'gsThisSet.MoveFirst
'For i = 0 To RecordCount - 1
'  If i = Combo2.ListIndex Then Exit For
'  gsThisSet.MoveNext
'Next
'gsThisSet.FindFirst "name='" + gsThisSet(Combo1.ListIndex).Fields("name").Value + "'"
'txtFields(1).Text = gsThisSet.Fields("number").Value
'txtFields(3).Text = gsThisSet.Fields("dou").Value
'Combo1.Text = Combo1.List(Combo1.ListIndex)
combo1.ListIndex = Combo2.ListIndex
Call Combo1_Click
End Sub


Private Sub Combo3_Change()

End Sub

Private Sub Form_Load()
china_english = CHINA
FirstDo
'Me.Top = 0
'Me.Height = 4800
'Me.Width = 6890
'Me.Left = 0
End Sub

Public Sub FirstDo()
Dim i, count As Integer
If china_english = True Then
  cmdAddNew.Caption = "添加记录"
  cmdEdit.Caption = "编辑记录"
  cmdDelete.Caption = "删除记录"
  cmdFind.Caption = "查询记录"
  cmdUpdate.Caption = "更新"
  cmdCancel.Caption = "取消"
  Frame1.Caption = ""
  Frame2.Caption = "移动记录"
  Label(0).Caption = "段号编号:"
  Label(1).Caption = "所投配方I:"
  Label(2).Caption = "所投配方II:"
  Label(3).Caption = "投料时间:"
  Label(4).Caption = "投料上升0:"
  Label(5).Caption = "转速:"
  Label(6).Caption = "压力级:"
  Label7.Caption = "记录位置:"
  Label(7).Caption = "时间:"
  Label(8).Caption = "能量:"
  Label(9).Caption = "控制关系:"
End If
Set gsDBName = OpenDatabase("c:\ljxt\ljxt.mdb")
'Set gsThisSet = gsDBName.OpenRecordset("tan_hei_table", dbOpenDynaset)
'Set gsThatSet = gsDBName.OpenRecordset("you_liao_table", dbOpenDynaset)
'gsThisSet.MoveLast
'count = gsThisSet.RecordCount
'gsThisSet.MoveFirst
'For i = 0 To count - 1
'  Combo2.AddItem "   " + gsThisSet.Fields("number").Value + Space(20) + gsThisSet.Fields("name").Value + Space(15 - Len(gsThisSet.Fields("name").Value)) + gsThisSet.Fields("dou").Value
'  Combo1.AddItem gsThisSet.Fields("name").Value
'  gsThisSet.MoveNext
'Next
'gsThisSet.MoveFirst
UseCombo1 = True
'gsFindField = ""
gsFindOpen = ""
gsFindExpror = ""
gsFindType = 0
End Sub

Private Sub txtFields_DblClick(Index As Integer)
Select Case Index
Case 1:
  UseCombo1 = True
  select_from.Show
Case 2:
  UseCombo1 = False
  select_from.Show
End Select
End Sub



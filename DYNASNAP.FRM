VERSION 4.00
Begin VB.Form frmDynaSnap 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4215
   ClientLeft      =   525
   ClientTop       =   1380
   ClientWidth     =   5295
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   4620
   HelpContextID   =   2016125
   Icon            =   "DYNASNAP.frx":0000
   KeyPreview      =   -1  'True
   Left            =   465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4196.911
   ScaleMode       =   0  'User
   ScaleWidth      =   5314.02
   Tag             =   "Recordset"
   Top             =   1035
   Width           =   5415
   Begin VB.CommandButton cmdlast 
      Caption         =   ">>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   17
      Top             =   840
      Width           =   372
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   16
      Top             =   840
      Width           =   372
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   372
   End
   Begin VB.PictureBox picViewButtons 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleMode       =   0  'User
      ScaleWidth      =   5298.159
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cl&ose"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2160
         TabIndex        =   14
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2160
         TabIndex        =   13
         Top             =   0
         Width           =   1092
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "D&elete"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdTrans 
         Caption         =   "&Chinese"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         Caption         =   "Label1"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   840
         Width           =   3495
      End
   End
   Begin VB.PictureBox picChangeButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   919.528
      ScaleMode       =   0  'User
      ScaleWidth      =   5719.056
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   960
         TabIndex        =   4
         Top             =   48
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2640
         TabIndex        =   5
         Top             =   48
         Width           =   1212
      End
   End
   Begin VB.PictureBox picFldHdr 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   15.009
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   15
      Begin VB.Label lblFieldHdr 
         Caption         =   "Field Name:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.VScrollBar vsbScrollBar 
      Height          =   1695
      LargeChange     =   3000
      Left            =   5040
      SmallChange     =   300
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picFields 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   372
      ScaleMode       =   0  'User
      ScaleWidth      =   4812
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4815
      Begin VB.TextBox txtFieldData 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblFieldName 
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmDynaSnap"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'Dim currdate As string
'Dim Filterredset As Recordset
'Dim FEditFlag As Integer
'Dim Faddflag As Integer
Dim Newflag As Integer


Private Sub cmdAdd_Click()
 'set the mode
 FAddNewFlag = True
 cmdFirst.Enabled = False
 cmdPrev.Enabled = False
 cmdNext.Enabled = False
 cmdLast.Enabled = False
 currdate = Date + time
 For i = 0 To mrecFormRecordset.Fields.count - 1
  If mrecFormRecordset(i).Type = dbDate Then
    frmDynaSnap.txtFieldData(i).Text = currdate
  Else
    frmDynaSnap.txtFieldData(i).Text = ""
  End If
 Next
 For i = 0 To mrecFormRecordset.Fields.count - 1
   If mrecFormRecordset(i).name = "number" Then
     Select Case gsRecordSet
      Case "tan_hei_table":
        frmDynaSnap.txtFieldData(i).Text = "T"
      Case "you_liao_table":
        frmDynaSnap.txtFieldData(i).Text = "Y"
      Case "jiao_liao_table":
        frmDynaSnap.txtFieldData(i).Text = "J"
      Case "lian_jiao_table":
        frmDynaSnap.txtFieldData(i).Text = "L"
      End Select
    End If
 Next
  mrecFormRecordset.AddNew
 ' mbAddNewFlag = True
  If mrecFormRecordset.RecordCount > 0 Then
    msBookMark = mrecFormRecordset.Bookmark
  Else
    msBookMark = ""
  End If

  picChangeButtons.Visible = True
  picViewButtons.Visible = False
  'hsclCurrRow.Enabled = False
  
  'ClearDataFields Me, mrecformrecordset.Fields.Count
  'txtFieldData(0).SetFocus
  'mbDataChanged = False
  Exit Sub

End Sub

Private Sub cmdCancel_Click()
picChangeButtons.Visible = False
   picViewButtons.Visible = True
'   hsclCurrRow.Enabled = True

   FEditFlag = False
   FAddNewFlag = False
   mrecFormRecordset.CancelUpdate
'   DBEngine.Idle dbFreeLocks
   If Newflag = True Then
   frmTables.Show
   Unload PaifangList
   Unload Me
   Else
 cmdFirst.Enabled = True
 cmdPrev.Enabled = True
 cmdNext.Enabled = True
 cmdLast.Enabled = True
 currdate = Date + time

   DisplayCurrSet ' Me, mrecformrecordset, mlNumRows, mbAddNewFlag
'   mbDataChanged = False
   End If

End Sub




Private Sub cmdClose_Click()
Unload PaifangList
Unload Me
frmmenu.Show
End Sub

Private Sub cmdDelete_Click()
mrecFormRecordset.Delete
If mrecFormRecordset.EOF = False Then
      mrecFormRecordset.MoveLast
DisplayCurrSet
'If mrecformrecordset.EOF Then
'  mrecformrecordset.MoveLast
End If

'mrecformrecordset.MovePrevious
End Sub

Private Sub cmdEdit_Click()
  FEditFlag = True
  cmdFirst.Enabled = False
  cmdPrev.Enabled = False
  cmdNext.Enabled = False
  cmdLast.Enabled = False
  currdate = Date + time

  mrecFormRecordset.Edit
  picChangeButtons.Visible = True
  picViewButtons.Visible = False
  'hsclCurrRow.Enabled = False
  currdate = Date + time
  For i = 0 To mrecFormRecordset.Fields.count - 1
  If mrecFormRecordset(i).Type = dbDate Then
    frmDynaSnap.txtFieldData(i).Text = currdate
  End If
 Next
  'ClearDataFields Me, mrecformrecordset.Fields.Count
  'txtFieldData(0).SetFocus
End Sub

Private Sub cmdFind_Click()
 'On Error GoTo FindErr
  
  Dim i As Integer
  Dim sBookMark As String
  Dim sTmp As String

  'load the column names into the find form
  If frmFind.lstFields.ListCount = 0 Then
    For i = 0 To mrecFormRecordset.Fields.count - 1
      frmFind.lstFields.AddItem Mid(lblFieldName(i).Caption, 1, Len(lblFieldName(i).Caption) - 1)
    Next
  End If

FindStart:

  'reset the flags
  gbFindFailed = False
  gbFromTableView = False
  mbNotFound = False
  Hide
  'MsgBar "Enter Search Parameters", False
  frmFind.Show vbModal
  'MsgBar "Searching for New Record", True
  If gbFindFailed = True Then   'find cancelled
    GoTo AfterWhile
  End If

  'SetHourglass

  'i = frmFind.lstFields.ListIndex
  sBookMark = mrecFormRecordset.Bookmark
  'search for the record
  If mrecFormRecordset(lstI).Type = dbText Or mrecFormRecordset(lstI).Type = dbMemo Then
    sTmp = AddBrackets((mrecFormRecordset(lstI).name)) & " " & gsFindOp & " '" & gsFindExpr & "'"
  Else
    sTmp = AddBrackets((mrecFormRecordset(lstI).name)) + gsFindOp + gsFindExpr
  End If
  Select Case gnFindType
    Case 0
      mrecFormRecordset.FindFirst sTmp
    Case 1
      mrecFormRecordset.FindNext sTmp
    Case 2
      mrecFormRecordset.FindPrevious sTmp
    Case 3
      mrecFormRecordset.FindLast sTmp
  End Select
  mbNotFound = mrecFormRecordset.NoMatch
  If mbNotFound = False Then
  DisplayCurrSet
  Exit Sub
  End If
AfterWhile:

  'Screen.MousePointer = vbDefault

  If gbFindFailed = True Then   'go back to original row
    MsgBox "Record Find Failed", 48
    mrecFormRecordset.Bookmark = sBookMark
  ElseIf mbNotFound Then
    Beep
    MsgBox "Record Not Found", 48
    mrecFormRecordset.Bookmark = sBookMark
    GoTo FindStart
  Else
    sBookMark = mrecFormRecordset.Bookmark  'save the new position
    'now we need to reposition the scrollbar to reflect the move
   ' If mlNumRows > 99 Then
   '   hsclCurrRow.VALUE = (mrecformrecordset.PercentPosition * mlNumRows) / 100 + 1
   ' Else
   '   hsclCurrRow.VALUE = mrecformrecordset.PercentPosition
   ' End If
    mrecFormRecordset.Bookmark = sBookMark
  End If

  DisplayCurrSet ' Record Me, mrecformrecordset, mlNumRows, mbAddNewFlag
  'mbDataChanged = False

  'MsgBar gsNULL_STR, False
  Exit Sub

FindErr:
'  Screen.MousePointer = vbDefault
'  If Err <> gnEOF_ERR Then
'    ShowError
    Exit Sub
'  Else
'    mbNotFound = True
'    Resume Next
'  End If

End Sub




Private Sub cmdFirst_Click()
mrecFormRecordset.MoveFirst
DisplayCurrSet
lblStatus.Caption = mrecFormRecordset.PercentPosition

End Sub

Private Sub cmdLast_Click()
mrecFormRecordset.MoveLast
DisplayCurrSet
lblStatus.Caption = mrecFormRecordset.PercentPosition
End Sub

Private Sub cmdNext_Click()
mrecFormRecordset.MoveNext
If mrecFormRecordset.EOF Then
  mrecFormRecordset.MoveLast
End If
DisplayCurrSet
lblStatus.Caption = mrecFormRecordset.PercentPosition
End Sub

Private Sub cmdPrev_Click()
mrecFormRecordset.MovePrevious
If mrecFormRecordset.BOF Then
  mrecFormRecordset.MoveFirst
End If
DisplayCurrSet
lblStatus.Caption = mrecFormRecordset.PercentPosition
End Sub

Private Sub cmdTrans_Click()
If china_english = False Then
  cmdAdd.Caption = "添加"
  cmdEdit.Caption = "编辑"
  cmdFind.Caption = "查找"
  cmdClose.Caption = "关闭"
  cmdTrans.Caption = "英文"
  cmdDelete.Caption = "删除"
  cmdUpdate.Caption = "更新"
  cmdCancel.Caption = "取消"
  lblFieldHdr.Caption = "字段名 "
  Dim fname As String
  For i = 0 To mrecFormRecordset.Fields.count - 1
     Select Case mrecFormRecordset(i).name
       Case "password": fname = "口令"
       Case "number": fname = "编号"
       Case "name": fname = "名称"
       Case "dou": fname = "斗号"
       Case "weight": fname = "所需重量"
       Case "fast_do": fname = "快慢提前量"
       Case "drop_do": fname = "点动提前量"
       Case "gon_cha": fname = "允许公差"
       Case "time": fname = "控制时间"
       Case "st": fname = "所"
       Case "w": fname = ""
       Case "g": fname = ""
       Case "Du1": fname = ""
       Case "Du2": fname = ""
       Case "tou": fname = ""
       Case "speed": fname = ""
       Case "pre": fname = ""
       Case "neg": fname = ""
       Case "ctr": fname = ""
       
       
       
     End Select
     lblFieldName(i).Caption = fname & ":"
     Next i
Else
  cmdAdd.Caption = "&Add"
  cmdEdit.Caption = "&Edit"
  cmdFind.Caption = "&Find"
  cmdClose.Caption = "Cl&ose"
  cmdTrans.Caption = "&China"
  cmdDelete.Caption = "&Delete"
  cmdUpdate.Caption = "&Update"
  cmdCancel.Caption = "&Cancel"
  lblFieldHdr.Caption = "FieldsName"
  For i = 0 To mrecFormRecordset.Fields.count - 1
    lblFieldName(i).Caption = mrecFormRecordset(i).name & ":"
  Next i
End If
End Sub


Private Sub cmdUpdate_Click()
On Error GoTo UpdateErr
 Dim nDelay As Long
  Dim nRetryCnt As Integer
retryupd: For i = 0 To mrecFormRecordset.Fields.count - 1
        If mrecFormRecordset(i).Type = dbMemo Then
          If mrecFormRecordset(i).FieldSize() < 50 Then
            mrecFormRecordset(i) = frmDynaSnap.txtFieldData(i).Text '= StripNonAscii(vFieldVal(mrecformrecordset(i)))
          Else
            mrecFormRecordset(i) = frmDynaSnap.txtFieldData(i).Text '= StripNonAscii(vFieldVal(mrecformrecordset(i).GetChunk(0, 50)))
          End If
    ElseIf mrecFormRecordset(i).Type = dbText Then
          mrecFormRecordset(i) = frmDynaSnap.txtFieldData(i).Text '= StripNonAscii(vFieldVal(mrecformrecordset(i)))
                
                Else: mrecFormRecordset(i) = vFieldVal(frmDynaSnap.txtFieldData(i).Text) '= vFieldVal(mrecformrecordset(i))
        End If
      Next
  'SetHourglass
  
  mrecFormRecordset.Update
  'If gbTransPending Then gbDBChanged = True

  'If mbAddNewFlag = True Then
'    mlNumRows = mlNumRows + 1
    'SetScrollBar
    'move to the new record
   'mrecformrecordset.Bookmark = mrecformrecordset.LastModified
  

  picChangeButtons.Visible = False
  picViewButtons.Visible = True
  'hsclCurrRow.Enabled = True
  FEditFlag = False
  FAddNewFlag = False
  cmdFirst.Enabled = True
  cmdPrev.Enabled = True
  cmdNext.Enabled = True
  cmdLast.Enabled = True
  currdate = Date + time

  'DisplayCurrentRecord Me, mrecformrecordset, mlNumRows, mbAddNewFlag
  DisplayCurrSet
  'mbDataChanged = False
  'DBEngine.Idle dbFreeLocks

  'Screen.MousePointer = vbDefault
  Exit Sub

UpdateErr:
  'check for locked error
  If Err = 3260 And nRetryCnt < gnMURetryCnt Then
    nRetryCnt = nRetryCnt + 1
    mrecFormRecordset.Bookmark = mrecFormRecordset.Bookmark   'Cancel the update
    DBEngine.Idle dbFreeLocks
    nDelay = Timer
    'Wait gnMUDelay seconds
    While Timer - nDelay < gnMUDelay
 '     'do nothing
    Wend
    Resume retryupd
  Else
'    ShowError
    Exit Sub
  End If



End Sub





Private Sub Form_Load()
Set mrecFormRecordset = gsDBName.OpenRecordset(gsRecordSet, dbOpenDynaset) 'frmDynaSnap.lstQueryDefs.Text)
'frmDynaSnap.Hide
Dim Fld As Field
Newflag = False
For Each Fld In mrecFormRecordset.Fields
   If Fld.name = "password" Then
     mrecFormRecordset.Filter = "[password]='" + password + "'"
     Set mrecFormRecordset = mrecFormRecordset.OpenRecordset()
   ElseIf Fld.name = "user_name" Then
     mrecFormRecordset.Filter = "[user_name]='" + user_name + "'"
     Set mrecFormRecordset = mrecFormRecordset.OpenRecordset()
   End If
Next
'mrecfromRecordset.perpentposition
Me.Caption = "Table:" & gsRecordSet + "  User_Name:" & user_name
Dim nFieldType As Integer
    lblFieldName(0).Top = lblFieldHdr.Top + 480&
    lblFieldName(0).Visible = True
    
    txtFieldData(0).Top = lblFieldHdr.Top + 480&
    txtFieldData(0).Visible = True
   nFieldType = mrecFormRecordset(0).Type
   txtFieldData(0).Width = GetFieldWidth(nFieldType)
   If nFieldType = dbText Then txtFieldData(0).MaxLength = mrecFormRecordset(0).Size
   txtFieldData(0).TabIndex = 0
   For i = 1 To mrecFormRecordset.Fields.count - 1
     picFields.Height = picFields.Height + 480&
     Load lblFieldName(i)
     lblFieldName(i).Top = lblFieldName(i - 1).Top + 480&
     lblFieldName(i).Visible = True
     Load txtFieldData(i)
     txtFieldData(i).Top = txtFieldData(i - 1).Top + 480&
     txtFieldData(i).Visible = True
     nFieldType = mrecFormRecordset.Fields(i).Type
     txtFieldData(i).Width = GetFieldWidth(nFieldType)
     If nFieldType = dbText Then txtFieldData(i).MaxLength = mrecFormRecordset(i).Size
     If nFieldType = dbDate Then txtFieldData(i).MaxLength = 2000
     txtFieldData(i).TabIndex = i
   Next

   'resize main window
   Me.Width = 4020 '5520
   If i <= 10 Then
     Me.Height = ((i + 1) * 480&) + 1690
   Else
     Me.Height = 5415 '4368
     Me.Width = Me.Width '+ 260
     'vsbScrollBar.Top = lblFieldHdr.Top + 480&
     vsbScrollBar.Visible = True
     vsbScrollBar.Min = 1553 '1080
     vsbScrollBar.Max = 1553 - (i * 480&) + 1690 '1080 - (i * 480&) + 2240
   End If
   SelectRecordSet
   'mrecFormRecordset.MoveLast
   counts = mrecFormRecordset.RecordCount
   'mrecFormRecordset.MoveFirst
   'display the field names
   For i = 0 To mrecFormRecordset.Fields.count - 1
     lblFieldName(i).Caption = mrecFormRecordset(i).name & ":"
   Next
   If (counts) <> 0 Then
    lblStatus.Caption = mrecFormRecordset.PercentPosition ' + " of " + Str(counts) ' Display current position.
      For i = 0 To mrecFormRecordset.Fields.count - 1
        If mrecFormRecordset(i).Type = dbMemo Then
           If mrecFormRecordset(i).FieldSize() < 50 Then
            frmDynaSnap.txtFieldData(i).Text = StripNonAscii(vFieldVal(mrecFormRecordset(i)))
           Else
            frmDynaSnap.txtFieldData(i).Text = StripNonAscii(vFieldVal(mrecFormRecordset(i).GetChunk(0, 50)))
           End If
       ElseIf mrecFormRecordset(i).Type = dbText Then
           frmDynaSnap.txtFieldData(i).Text = StripNonAscii(vFieldVal(mrecFormRecordset(i)))
          Else
            frmDynaSnap.txtFieldData(i).Text = vFieldVal(mrecFormRecordset(i))
          End If
      Next
   Else:
   lblStatus.Caption = "没发现"
   For i = 0 To (mrecFormRecordset.Fields.count - 1)
   frmDynaSnap.txtFieldData(i).Text = ""
   Next
   Newflag = True
   Call cmdAdd_Click
   End If
   Call cmdTrans_Click
End Sub

Function GetFieldWidth(rType As Integer)
  Select Case rType
    Case dbBoolean
      GetFieldWidth = 850
    Case dbByte
      GetFieldWidth = 650
    Case dbInteger
      GetFieldWidth = 900
    Case dbLong
      GetFieldWidth = 1100
    Case dbCurrency
      GetFieldWidth = 1800
    Case dbSingle
      GetFieldWidth = 1800
    Case dbDouble
      GetFieldWidth = 2200
    Case dbDate
      GetFieldWidth = 2000
    Case dbText
      GetFieldWidth = 3250
    Case dbLongBinary
      GetFieldWidth = 3250
    Case dbMemo
      GetFieldWidth = 3250
    Case Else
      GetFieldWidth = 3250
  End Select

End Function


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


Private Sub hsclCurrRow_Change()
mrecFormRecordset.MoveNext
For i = 0 To mrecFormRecordset.Fields.count - 1
        If mrecFormRecordset(i).Type = dbMemo Then
          If mrecFormRecordset(i).FieldSize() < 50 Then
            frmDynaSnap.txtFieldData(i).Text = StripNonAscii(vFieldVal(mrecFormRecordset(i)))
          Else
            frmDynaSnap.txtFieldData(i).Text = StripNonAscii(vFieldVal(mrecFormRecordset(i).GetChunk(0, 50)))
          End If
    ElseIf mrecFormRecordset(i).Type = dbText Then
          frmDynaSnap.txtFieldData(i).Text = StripNonAscii(vFieldVal(mrecFormRecordset(i)))
        Else
          frmDynaSnap.txtFieldData(i).Text = vFieldVal(mrecFormRecordset(i))
        End If
      Next

End Sub








Private Sub Form_Resize()
  On Error Resume Next

  Dim nHeight As Integer
  Dim i As Integer
  Dim nTotWidth As Integer
  Const nHeightFactor = 1420

  'If WindowState <> 1 Then   'not minimized
   ' MsgBar "Resizing Form", True
    'make sure the form is lined up on a field
    nHeight = Height
    If (nHeight - nHeightFactor) Mod 380& <> 0 Then
      Me.Height = ((nHeight - nHeightFactor) \ 380&) * 380& + nHeightFactor
    End If
    'resize the status bar
    'picMoveButtons.Top = Me.Height - 650
    Label1.Top = Me.Height - 1570
    'resize the scrollbar
   ' vsbScrollBar.Height = picMoveButtons.Top - (picViewButtons.Top - picFldHdr.Height) - 1320
    vsbScrollBar.Height = DBGrid1.Top - (picViewButtons.Top - picFldHdr.Height) - 1320
    vsbScrollBar.Left = Me.Width - 360
    If mrecFormRecordset.Fields.count > 10 Then
      picFields.Width = Me.Width - 260
      nTotWidth = vsbScrollBar.Left - 20
    Else
      picFields.Width = Me.Width - 20
      nTotWidth = Me.Width - 50
    End If
    picFldHdr.Width = Me.Width - 20
    'widen the fields if possible
    For i = 0 To mrecFormRecordset.Fields.count - 1
      lblFieldName(i).Width = 0.3 * nTotWidth
      txtFieldData(i).Left = lblFieldName(i).Width + 20
      If mrecFormRecordset(i).Type = dbText Or mrecFormRecordset(i).Type = dbMemo Then
        txtFieldData(i).Width = 0.7 * nTotWidth - 250
      End If
    Next
    lblFieldValue.Left = txtFieldData(0).Left
    'hsclCurrRow.Width = picMoveButtons.Width \ 2
    'lblStatus.Width = picMoveButtons.Width \ 2
    'lblStatus.Left = hsclCurrRow.Width + 10
  'End If
End Sub

Private Sub txtFieldData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 34 And vsbScrollBar.Visible = True Then
    'pagedown with > 10 fields
    vsbScrollBar.Value = vsbScrollBar.Value - 3000

  ElseIf KeyCode = 33 And vsbScrollBar.Visible = True Then
    'pageup with > 10 fields
    vsbScrollBar.Value = vsbScrollBar.Value + 3000

  End If
End Sub

Private Sub txtFieldData_KeyPress(Index As Integer, KeyAscii As Integer)
'only allow return when in edit of add mode
  If FEditFlag = True Or Faddflag = True Then
    If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{Tab}"
    End If

  'throw away the keystrokes if not in add or edit mode
  ElseIf FEditFlag = False And Faddflag = False Then
    KeyAscii = 7
  End If

End Sub


Private Sub vsbScrollBar_Change()
Dim nTop As Integer

  nTop = vsbScrollBar.Value
  If (nTop - 1553) Mod 480& = 0 Then
    picFields.Top = nTop
  Else
    picFields.Top = ((nTop - 1553) \ 480&) * 480& + 1670
  End If
End Sub



Public Sub SelectRecordSet()
  Select Case gsRecordSet
    Case "tan_hei_table":
      gsBaseRecordSet = "cai_liao_table"
    Case "you_liao_table":
      gsBaseRecordSet = "cai_liao_table"
    Case "jiao_liao_table":
      gsBaseRecordSet = "cai_liao_table"
    Case "lian_jiao_table":
      gsBaseRecordSet = "cai_liao_table"
    'Case "":
    '  gsBaseRecordSet = ""
default: gsBaseRecordSet = ""
         Exit Sub
  End Select
  PaifangList.Show
End Sub

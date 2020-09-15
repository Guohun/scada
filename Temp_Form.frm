VERSION 4.00
Begin VB.Form Temp_Form 
   Caption         =   "temp"
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   1485
   ClientWidth     =   9135
   Height          =   4905
   Left            =   90
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   9135
   Top             =   1140
   Width           =   9255
   Begin VB.CommandButton Command3 
      Caption         =   "write_read"
      Height          =   615
      Left            =   1200
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox read_ban_text 
      Height          =   855
      Left            =   5160
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox read_dou_text 
      Height          =   855
      Left            =   5400
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "read"
      Height          =   855
      Left            =   6240
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "write"
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox write_dou_text 
      Height          =   975
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox write_ban_text 
      Height          =   975
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   2880
   End
   Begin VB.Label Label4 
      Caption         =   "read_ban"
      Height          =   735
      Left            =   3960
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "read_dou"
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "write_ban"
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "write_dou"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Temp_Form"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Static Function Speak_Error(ByVal error_num As Integer) As Integer
  Dim db As Database
  Dim tb As TableDefs
  Set db = OpenDatabase("c:\ljxt\ljxt.mdb")
  Set tb = db.TableDefs("error_mnu")
  If (error_num < 0) And (error_num > td.Field.count) Then
    MsgBox "ghghhjjjjjj"
    Speak_Error = 0
    Else
       Speak_Error = 1
    End If
        Exit Function
  
  
End Function

Static Function kill_error(ByVal error_num As Integer) As Integer
  Dim db As Database
  Dim tb As TableDefs
  Set db = OpenDatabase("c:\ljxt\ljxt.mdb")
  Set tb = db.TableDefs("error_mnu")
    If (error_num < 0) And (error_num > td.Field.count) Then
        MsgBox "ghghhjjjjjj"
        kill_error = 0
    Else
      
      kill_error = 1
    
        End If
        Exit Function
End Function

 Private Sub ChangeButtons()
  cmdAdd.visible = False
  cmdEdit.visible = False
'  cmdFind.Visible = False
  cmdDel.visible = False
  cmdUpdate.visible = True
  cmdCancel.visible = True
'  txtFields(2).Visible = False
  Frame2.visible = False
  cmdFirst.visible = False
  cmdPrev.visible = False
  cmdNext.visible = False
  cmdLast.visible = False
'  Combo2.Visible = True
End Sub

Private Sub RestoreButtons()
  cmdAdd.visible = True
  cmdEdit.visible = True
  cmdDel.visible = True
'  cmdFind.Visible = True
  cmdUpdate.visible = False
  cmdCancel.visible = False
  Frame2.visible = True
  cmdFirst.visible = True
  cmdPrev.visible = True
  cmdNext.visible = True
  cmdLast.visible = True

End Sub


Private Sub FirstDo()
    

 If china_english = CHINA Then
     Caption = " 报警表录入目录 "
     
      cmdAdd.Caption = "添加记录 "
     cmdEdit.Caption = "编辑记录"
     cmdDel.Caption = "删除记录"
'     cmdFind.Caption = "查询记录"
     cmdUpdate.Caption = "更新"
     cmdCancel.Caption = "取消"
     Frame2.Caption = " 移动记录"
     error_name_label.Caption = "   错误名 "
     china_speak_label.Caption = " 中文声音"
     china_prompt_label.Caption = " 中文提示"
     eng_speak_label.Caption = "英文声音"
     eng_prompt_label.Caption = " 英文提示"
     china_command.Caption = " 中文声音处理"
     eng_command.Caption = " 英文声音处理"
     
'     Text2.Text = "   起动点"
 '    Text3.Text = "   检测点"
  '   Text5.Text = "   实际启动"
   '  Text6.Text = "   实际检测"
   '  Text4.Text = "   延时"
   '  Text7.Text = "   启动点"
   '  Text8.Text = "    检测点"
   '  Text11.Text = "   实际检测"
   '  Text10.Text = "   实际启动"
    ' Text9.Text = "    延时"
  End If
  RestoreButtons
End Sub



Private Sub chinese_Change()

End Sub

Private Sub china_command_Click()
        Dim msg  As String
        OLE1.SizeMode = 2   ' Autosize.
        OLE1.CreateEmbed "", "soundrec"
        OLE1.DoVerb -3
'        OLE1.DoVerb -2
        If OLE1.AppIsRunning Then
        ' msg = OLE1.DataText
        ' Update the object.
        OLE1.Update
    Else
        MsgBox "wave  isn't active."
    End If
End Sub



Private Sub cmdAdd_Click()
 ChangeButtons
 lian_liao_Data.Recordset.AddNew
' lian_liao_Data.Recordset.Fields("Password").Value = Password
   If Text12.Text = "2" Then
        Text12.Text = "2"
   Else
        Text12.Text = "1"
   End If
 
 If lian_liao_Data.Recordset.RecordCount > 0 Then
   msTBookMark = lian_liao_Data.Recordset.Bookmark
 Else
   msTBookMark = ""
  End If

End Sub

Private Sub cmdCancel_Click()
    lian_liao_Data.Recordset.CancelUpdate
    RestoreButtons
End Sub

Private Sub cmdDel_Click()
If china_english = CHINA Then
  If MsgBox("真的要删除该记录吗？", 4, "记录删除") = 6 Then
  lian_liao_Data.Recordset.Delete
'  Call cmdPrev_Click
Call cmdPrev_Click
  End If
Else
  If MsgBox("Really Want To Delete Current Record?", 4, "Record Delete") = 6 Then lian_liao_Data.Recordset.Delete
End If

End Sub

Private Sub cmdEdit_Click()
Dim LIndex As Integer
lian_liao_Data.RecordsetType = vbRSTypeDynaset
lian_liao_Data.Recordset.Edit
ChangeButtons
'LIndex = 0
'If txtFields(2).Text <> "" Then
'Do
' If (txtFields(2).Text) = combo1.List(LIndex) Then Exit Do
' LIndex = LIndex + 1
'Loop
'End If
'combo1.ListIndex = LIndex

End Sub

Private Sub cmdFirst_Click()
    If lian_liao_Data.Recordset.RecordCount > 1 Then
      lian_liao_Data.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If lian_liao_Data.Recordset.RecordCount > 0 Then
        lian_liao_Data.Recordset.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
cmdNext.enabled = False
If lian_liao_Data.Recordset.RecordCount > 1 And Not lian_liao_Data.Recordset.EOF Then
    lian_liao_Data.Recordset.MoveNext
    If lian_liao_Data.Recordset.EOF Then
            lian_liao_Data.Recordset.MovePrevious
'        Resume
    End If
End If
cmdNext.enabled = True
End Sub

Private Sub cmdPrev_Click()
cmdPrev.enabled = False
If lian_liao_Data.Recordset.RecordCount > 1 And Not lian_liao_Data.Recordset.BOF Then
    lian_liao_Data.Recordset.MovePrevious
    If lian_liao_Data.Recordset.BOF Then
            lian_liao_Data.Recordset.MoveNext
'        Resume
    End If
End If
cmdPrev.enabled = True
End Sub

Private Sub cmdUpdate_Click()
      lian_liao_Data.Recordset.Update
      RestoreButtons
End Sub

Private Sub lian_liao_Data_Reposition()
If china_english = CHINA Then
  If Not lian_liao_Data.Recordset.EOF And Not lian_liao_Data.Recordset.BOF Then
        record_prompt.Caption = "当前记录:" & (lian_liao_Data.Recordset.AbsolutePosition + 1) & " 之 " & lian_liao_Data.Recordset.RecordCount
Else
        record_prompt.Caption = "没有记录:"
  End If
Else

End If

End Sub

Private Sub eng_command_Click()
        OLE2.SizeMode = 2   ' Autosize.
        OLE2.CreateEmbed "", "soundrec"
        OLE2.DoVerb -3
'        OLE1.DoVerb -2
        If OLE2.AppIsRunning Then
        ' msg = OLE1.DataText
        ' Update the object.
        OLE2.Update
    Else
        MsgBox "wave  isn't active."
    End If
End Sub

Private Sub exit_command_Click()
        Unload Me
End Sub

Private Sub Command1_Click()
        Timer1.Interval = 10
        Timer2.Interval = 0
End Sub

Private Sub Command2_Click()
        Timer1.Interval = 0
        Timer2.Interval = 10
End Sub

Private Sub Command3_Click()
        Timer1.Interval = 10
        Timer2.Interval = 10

End Sub

Private Sub Form_Load()
    Call can_read("12", "1", 1, 1, 1)
End Sub




Private Sub txtchina_Change()

End Sub


Private Sub Timer1_Timer()
    Static x As Integer
    Dim y  As Single
    Dim Z  As Single

        x = x + 1
        y = x
        Z = x
         Call write_ban(1, 1, y, Z)
          Call write_select_dou(1, 1, x)
            write_dou_text.Text = CStr(x)
            write_ban_text.Text = CStr(y)
        If x > 100 Then x = 0
        
End Sub


Private Sub Timer2_Timer()
   Dim x As Integer
    Dim y As Single
        Dim Z  As Single
        Call read_ban(1, 1, y, Z)
         read_ban_text.Text = CStr(y)
          read_dou_text.Text = CStr(Z)
        
        Call read_select_dou(1, 1, x)
           If x > 100 Then x = 0

End Sub



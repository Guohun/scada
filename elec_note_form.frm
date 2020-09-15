VERSION 5.00
Begin VB.Form elec_node_form 
   Caption         =   "elec_node_form"
   ClientHeight    =   5688
   ClientLeft      =   744
   ClientTop       =   756
   ClientWidth     =   8508
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5688
   ScaleWidth      =   8508
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
      Begin VB.CommandButton Exit_Command 
         Caption         =   "&Exit"
         Height          =   500
         Left            =   120
         TabIndex        =   3
         Top             =   4320
         Width           =   800
      End
      Begin VB.CommandButton Save_Command 
         Caption         =   "&Save"
         Height          =   500
         Left            =   120
         TabIndex        =   2
         Top             =   3480
         Width           =   800
      End
      Begin VB.PictureBox Not_Image 
         Height          =   612
         Left            =   120
         Picture         =   "elec_note_form.frx":0000
         ScaleHeight     =   564
         ScaleWidth      =   564
         TabIndex        =   1
         Top             =   240
         Width           =   612
      End
      Begin VB.PictureBox Light_Image 
         ForeColor       =   &H0000FF00&
         Height          =   612
         Left            =   120
         ScaleHeight     =   564
         ScaleWidth      =   564
         TabIndex        =   5
         Top             =   2280
         Width           =   612
      End
      Begin VB.PictureBox Check_Image 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H0000FF00&
         Height          =   732
         Left            =   120
         ScaleHeight     =   684
         ScaleWidth      =   564
         TabIndex        =   4
         Top             =   1080
         Width           =   612
      End
   End
   Begin VB.PictureBox Image1 
      BackColor       =   &H00FFFFFF&
      Height          =   732
      Index           =   0
      Left            =   1320
      ScaleHeight     =   684
      ScaleWidth      =   564
      TabIndex        =   6
      Top             =   600
      Width           =   612
   End
End
Attribute VB_Name = "elec_node_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim data As Database
Dim rec As Recordset
Dim gcount As Integer


Const LIGHT_DOING = 0
Const CHECK_DOING = 1
Const NOTHING_DOING = 2
Public Place_Flag    As Byte
Public Edit_Flag   As Byte


Public Current_node_name     As Integer
Public Current_china_str     As String
Public Current_eng_str     As String
Public Current_name         As String
Public Current_row         As Integer
Public Current_col         As String
Public Current_Listindex         As Integer

Dim Move_Now    As Boolean
Public change_flag    As Boolean




'
Private Function First_Do()
Dim i   As Integer
Set data = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.mdb")
Set rec = data.OpenRecordset("select a.name,row,col,a.lenght,china_str,eng_str, b.alias,b.node_name,b.in_out from  elec_disp a ,node_table   b where b.alias=a.name ", dbOpenDynaset, dbInconsistent)
If rec.RecordCount > 0 Then
    rec.MoveLast
    gcount = rec.RecordCount
    For i = 1 To gcount - 1
                Load IMAGE1(i)
     Next i
                rec.MoveFirst
                For i = 0 To gcount - 1
                        IMAGE1(i).NodeName = rec.Fields("node_name").Value  '节点名
                        IMAGE1(i).gname = rec.Fields("name").Value
                        IMAGE1(i).Top = rec.Fields("row").Value
                        IMAGE1(i).Left = rec.Fields("col").Value
                       
                         IMAGE1(i).ENGLISH = vFieldVal(rec.Fields("eng_str").Value)
                         IMAGE1(i).CHINA = vFieldVal(rec.Fields("china_str").Value)
                        IMAGE1(i).Style = rec.Fields("in_out").Value
                        If china_english = CHINA Then
                                IMAGE1(i).Caption = IMAGE1(i).CHINA 'gchina_str(i)
                        Else
                                IMAGE1(i).Caption = IMAGE1(i).ENGLISH 'geng_str(i)
                        End If
                                IMAGE1(i).Visible = True
                        IMAGE1(i).Height = light_IMAGE.Height
                        IMAGE1(i).Width = check_image.Width
                         rec.MoveNext
                Next i
Else
        gcount = 1
End If
End Function

'0---  light
'1-----check
'2 ---nothing
Private Sub Elec_In(ByVal Place_x As Integer)
        Select Case Place_x
                 Case CHECK_DOING
                    For i% = 0 To gcount - 1
                    If IMAGE1(i%).Style = CHECK_DOING Then        '  阀
                        IMAGE1(i%).Visible = True
                    Else
                        IMAGE1(i%).Visible = False
                   End If
                   Next i%
                Case LIGHT_DOING
                   For i% = 0 To gcount - 1
                    If IMAGE1(i%).Style = LIGHT_DOING Then        '  阀
                            IMAGE1(i%).Visible = True
                    Else
                        IMAGE1(i%).Visible = False
                   End If
                   Next i%
            Case Else
                  For i% = 0 To gcount - 1
                            IMAGE1(i%).Visible = True
                   Next i%
            End Select
End Sub


Private Sub check_image_Click()
        Init_Color
        check_image.ForeColor = RGB(244, 244, 244)
        Place_Flag = CHECK_DOING
        Elec_In (CHECK_DOING)
End Sub






Private Sub exit_command_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Width = MAIN_MDI.Width  ' Set width of form.
   Height = MAIN_MDI.Height * 3 / 4 ' Set height of form.
     Left = 0
    Top = 0
        Place_Flag = NOTHING_DOING
        Edit_Flag = NOTHING_DOING
       Call Add_Win(Hwnd)

        change_flag = UNCHANGE
       Call First_Do
     Call Elec_In(NOTHING_DOING)

End Sub





Public Function Add_list()
'   node_List.Clear
 '   For i = 1 To gcount - 1
  '      node_List.AddItem g_node_name(i)
   ' Next i
    'node_List.ListIndex = 0
End Function



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        If Place_Flag = LIGHT_DOING Then
                If gcount > 50 Then
                        MsgBox "error load>50个信号  "
                End If
                Current_node_name = -1
                Current_name = ""
                Current_china_str = ""
                Current_eng_str = ""
                Current_row = Y
                Current_col = x
                elec_place_Form.Show
        End If
        
        If Place_Check = LIGHT_DOING Then
                If gcount > 50 Then
                        MsgBox "error load>50个信号  "
                End If
                Current_node_name = -1
                Current_name = ""
                Current_china_str = ""
                Current_eng_str = ""
                Current_row = Y
                Current_col = x
                elec_place_Form.Show
        End If

End Sub

Private Sub Form_Paint()
If china_english = CHINA Then
    check_image.Caption = "阀"
    light_IMAGE.Caption = "灯"
    exit_command.Caption = "&E退出"
    Save_Command.Caption = "&S保存"
Else
    on_label.Caption = "ON"
    off_label.Caption = "FF"
    exit_command.Caption = "&Exit"
    Save_Command.Caption = "&Save"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Answer As Integer
If change_flag = UNCHANGE Then
'        Unload Me
        Cancel = 0
        Exit Sub
End If
  Answer = MsgBox(" 数据已修改  确定退出吗？", 48 + vbYesNo, "提示")
  If Answer = vbYes Then
'        Unload Me
    Cancel = 0
  Else
    Cancel = 1
  End If

End Sub

Private Sub Form_Terminate()
    Call Del_Win(Hwnd)
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
                Move_Now = True
    Else                            'vbRightButton
                
    End If
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'   if button=  then
If (Button = vbLeftButton) Then
  If Move_Now = True Then
          change_flag = CHANGE
          IMAGE1(Index).Move x + IMAGE1(Index).Left, Y + IMAGE1(Index).Top
   End If
End If
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
     If Move_Now = True Then
        Move_Now = False
     End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    For i = 1 To gcount - 1
            Unload IMAGE1(i)
    Next i
    rec.Close
    data.Close
    Call Del_Win(Hwnd)

End Sub



Private Sub Image1_DblClick(Index As Integer)
            Edit_Flag = IMAGE1(Index).Style

              Current_Listindex = Index
              Current_node_name = IMAGE1(Current_Listindex).NodeName
              Current_name = IMAGE1(Current_Listindex).gname
              Current_china_str = IMAGE1(Current_Listindex).CHINA
              Current_eng_str = IMAGE1(Current_Listindex).ENGLISH
              elec_place_Form.Show
        
End Sub


Private Sub Image1_KeyPress(Index As Integer, KeyAscii As Integer)
' MsgBox KeyAscii
'   If KeyAscii = 13 Then
 '       If light_status(Index) = False Then
   '                Image1(i).Style = LIGHT_DOING
  '                 light_status(Index) = True
    '    Else
      '             Image1(i).Style = CHECK_DOING
     '              light_status(Index) = False
       ' End If
'End If
End Sub





Public Function Add_Light(ByVal row As Integer, ByVal col As Integer)

                Load IMAGE1(gcount)
                IMAGE1(gcount).CHINA = Current_china_str
                IMAGE1(gcount).ENGLISH = Current_eng_str
                IMAGE1(gcount).gname = Current_name
                
                IMAGE1(gcount).NodeName = Current_node_name
                IMAGE1(gcount).Visible = True
                IMAGE1(gcount).Style = LIGHT_DOING
                If china_english = CHINA Then
                        IMAGE1(gcount).Caption = Current_china_str
                Else
                        IMAGE1(gcount).Caption = Current_eng_str
                End If
                IMAGE1(gcount).Left = col
                IMAGE1(gcount).Top = row
                IMAGE1(gcount).Height = light_IMAGE.Height
                IMAGE1(gcount).Width = light_IMAGE.Width
            gcount = gcount + 1
                   
End Function




Private Sub light_IMAGE_Click()
        Init_Color
        light_IMAGE.ForeColor = RGB(244, 244, 244)
        Place_Flag = LIGHT_DOING
        Elec_In (LIGHT_DOING)
End Sub







Private Sub Init_Color()
        Not_Image.ForeColor = &HFF00&
        light_IMAGE.ForeColor = &HFF00&
        check_image.ForeColor = &HFF00&
End Sub



Private Sub node_List_Click()

End Sub

Private Sub Not_Image_Click()
        Init_Color
        Not_Image.BackColor = QBColor(4)
        Place_Flag = NOTHING_DOING
        Elec_In (NOTHING_DOING)
End Sub



Public Function Add_Check(ByVal row As Integer, ByVal col As Integer)

                Load IMAGE1(gcount)

                IMAGE1(gcount).CHINA = Current_china_str
                IMAGE1(gcount).ENGLISH = Current_eng_str
                IMAGE1(gcount).gname = Current_name
                
                IMAGE1(gcount).NodeName = Current_node_name
                IMAGE1(gcount).Visible = True
                IMAGE1(gcount).Style = LIGHT_DOING
                If china_english = CHINA Then
                        IMAGE1(gcount).Caption = Current_china_str
                Else
                        IMAGE1(gcount).Caption = Current_eng_str
                End If
                
                IMAGE1(gcount).Left = col
                IMAGE1(gcount).Top = row
                IMAGE1(gcount).Height = check_image.Height
                IMAGE1(gcount).Width = check_image.Width
                  
            gcount = gcount + 1
                   

End Function

Public Function Edit_Light()

    
                IMAGE1(Current_Listindex).CHINA = Current_china_str
                IMAGE1(Current_Listindex).ENGLISH = Current_eng_str
                IMAGE1(Current_Listindex).gname = Current_name
                IMAGE1(Current_Listindex).NodeName = Current_node_name
                
                
                IMAGE1(Current_Listindex).Visible = True
                
                
                If china_english = CHINA Then
                        IMAGE1(Current_Listindex).Caption = IMAGE1(Current_Listindex).CHINA
                Else
                        IMAGE1(Current_Listindex).Caption = IMAGE1(Current_Listindex).ENGLISH
                End If
        
End Function
Public Function Edit_Check()
                IMAGE1(Current_Listindex).CHINA = Current_china_str
                IMAGE1(Current_Listindex).ENGLISH = Current_eng_str
                IMAGE1(Current_Listindex).gname = Current_name
                IMAGE1(Current_Listindex).NodeName = Current_node_name
                
                
                IMAGE1(Current_Listindex).Visible = True
                
                
                If china_english = CHINA Then
                        IMAGE1(Current_Listindex).Caption = IMAGE1(Current_Listindex).CHINA
                Else
                        IMAGE1(Current_Listindex).Caption = IMAGE1(Current_Listindex).ENGLISH
                End If
End Function


Public Function Del_Elec()
Dim i%, temp           As Byte
                temp = IMAGE1(Current_Listindex).Style
                For i% = Current_Listindex To gcount - 2
                        IMAGE1(i%).gname = IMAGE1(i% + 1).gname
                        IMAGE1(i%).Style = IMAGE1(i% + 1).Style
                        IMAGE1(i%).ENGLISH = IMAGE1(i% + 1).ENGLISH
                        IMAGE1(i%).CHINA = IMAGE1(i% + 1).CHINA
                        IMAGE1(i%).Status = IMAGE1(i% + 1).Status
                        IMAGE1(i%).NodeName = IMAGE1(i% + 1).NodeName
                        IMAGE1(i%).Relation = IMAGE1(i% + 1).Relation
                        IMAGE1(i%).Caption = IMAGE1(i% + 1).Caption
                        IMAGE1(i%).Left = IMAGE1(i% + 1).Left
                        IMAGE1(i%).Top = IMAGE1(i% + 1).Top
                Next i%
                gcount = gcount - 1
                IMAGE1(gcount).Visible = False
                If gcount > 0 Then
                            Unload IMAGE1(gcount)
                End If
                Me.Cls
                Elec_In (temp)
End Function


Private Function Save_Node()
    Dim RecsChanged  As Integer
    Dim SQL As String
    If rec.RecordCount > 0 Then
    rec.MoveFirst
    Do While Not rec.EOF
            rec.Delete
            rec.MoveNext
    Loop
    
    End If
'    For i = 1 To gcount - 1
 '       SQL = "DELETE FROM elec_disp  where  name like '" & gname(i) & "' "
  '      Debug.Print SQL
  '      data.Execute SQL, dbDenyWrite
  '  Next i
'    RecsChanged = data .ExecuteSQL MySQL
    For i = 1 To gcount - 1
         rec.AddNew  ' Enable editing.
         rec.Fields("name").Value = IMAGE1(i).gname
         rec.Fields("alias").Value = IMAGE1(i).gname
          rec.Fields("row").Value = IMAGE1(i).Top                     'grow(i)
          rec.Fields("col").Value = IMAGE1(i).Left
         rec.Fields("node_name").Value = IMAGE1(i).NodeName '节点名
          rec.Fields("china_str").Value = IMAGE1(i).CHINA
          rec.Fields("in_out").Value = IMAGE1(i).Style
          rec.Fields("eng_str").Value = IMAGE1(i).ENGLISH
          rec.Update    ' Save changes.
Next i
End Function

Private Sub Save_Command_Click()
  change_flag = UNCHANGE
  Call Save_Node
  Call Init_Elec

End Sub



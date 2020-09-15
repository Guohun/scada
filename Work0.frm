VERSION 5.00
Begin VB.Form elec_change 
   Caption         =   "Form1"
   ClientHeight    =   6012
   ClientLeft      =   192
   ClientTop       =   468
   ClientWidth     =   8964
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6012
   ScaleWidth      =   8964
   Begin VB.CommandButton Change_Flag 
      Caption         =   "Change"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Exit_command 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6240
      Top             =   3960
   End
   Begin VB.PictureBox Image1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   1080
      ScaleHeight     =   684
      ScaleWidth      =   1164
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "elec_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim data As Database
Dim rec As Recordset
Dim tim 'As idoit 'Integer
Dim gcount As Integer
Dim gname(50) As String
Dim grow(50) As Integer
Dim gcol(50) As Integer
Dim geof(50) As Integer
Dim glenght(50) As Integer
Dim gtype(50) As Integer
Dim gwei(50) As Integer
Dim gchina_str(50) As String
Dim gchina_dat(50) As String
Dim geng_str(50) As String
Dim geng_dat(50) As String
Dim light_status(50) As Boolean
Dim Move_Now    As Boolean
Dim ChangeFlag   As Boolean












Private Sub change_flag_Click()
    If ChangeFlag = UNCHANGE Then
            Timer1.Interval = 0
            ChangeFlag = CHANGE
    Else
                Timer1.Interval = 20
                ChangeFlag = UNCHANGE
    End If
    
End Sub


Private Sub exit_command_Click()
    Unload Me
End Sub


Private Sub Form_Load()
 On Error Resume Next
 Call Add_Win(Hwnd)
    Call Init_Font(72, "宋体")
  Width = MAIN_MDI.Width  ' Set width of form.
   Height = MAIN_MDI.Height - 600 ' Set height of form.
     Left = 0
    Top = 0
If china_english = CHINA Then
        exit_command.Caption = "&E退出"
        change_flag.Caption = "&C修改"
Else

End If
tim = False
Set data = OpenDatabase("c:\ljxt\ljxt.mdb")
 Set rec = data.OpenRecordset("select * from  elec_disp a ,node_table   b where b.alias=a.name and b.in_out=1 ")
 If rec.RecordCount > 0 Then
    rec.MoveLast
    gcount = rec.RecordCount
    For i = 1 To gcount
            Load IMAGE1(i)
'            light_status(i) = False
    Next i
    rec.MoveFirst
    Elec_In
Else
    gcount = 1
End If
End Sub



Public Sub Elec_In()
   rec.MoveFirst
    For i = 0 To gcount - 1
        gname(i) = vFieldVal(rec.Fields("name").Value)
        grow(i) = rec.Fields("row").Value
        gcol(i) = rec.Fields("col").Value
        geof(i) = vFieldVal(rec.Fields("eof").Value)
        glenght(i) = vFieldVal(rec.Fields("lenght").Value)
        gtype(i) = vFieldVal(rec.Fields("type").Value)
        gwei(i) = vFieldVal(rec.Fields("wei").Value)
        gchina_str(i) = vFieldVal(rec.Fields("china_str").Value)
        gchina_dat(i) = vFieldVal(rec.Fields("china_dat").Value)
        geng_str(i) = vFieldVal(rec.Fields("eng_str").Value)
        geng_dat(i) = vFieldVal(rec.Fields("eng_dat").Value)
            If china_english = CHINA Then
'                Label1(i).Caption = gchina_str(i)
                 IMAGE1(i).Caption = gchina_str(i)
            Else
'                Label1(i).Caption = geng_str(i)
                IMAGE1(i).Caption = geng_str(i)
            End If
            IMAGE1(i).Visible = True
            IMAGE1(i).Status = 0
'            Image1(i).height = off_IMAGE.height
 '           Image1(i).width = off_IMAGE.width
            IMAGE1(i).Left = gcol(i)
            IMAGE1(i).Top = grow(i)
            
           rec.MoveNext
Next i
    
      

End Sub

Private Sub Form_Unload(Cancel As Integer)
    For i = 1 To gcount
            Unload IMAGE1(i)
    Next i
    rec.Close
    data.Close
Call Del_Win(Hwnd)

End Sub



Private Sub Image1_DblClick(Index As Integer)
        If light_status(Index) = False Then
                   'Image1(Index).Picture = on_image.Picture
                   IMAGE1(Index).Status = 1
                   light_status(Index) = True
        Else
                   IMAGE1(Index).Status = 0
'                   Image1(Index).Picture = off_IMAGE.Picture
                   light_status(Index) = False
        End If
End Sub


Private Sub Image1_GotFocus(Index As Integer)
    If ChangeFlag = CHANGE Then
            IMAGE1(Index).BackColor = RGB(100, 100, 100)
    End If
End Sub


Private Sub Image1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
        If light_status(Index) = False Then
                   IMAGE1(i).Status = 1
        Else
                IMAGE1(i).Status = 0
        End If
End If
End Sub


Private Sub Image1_LostFocus(Index As Integer)
    If ChangeFlag = CHANGE Then
            IMAGE1(Index).BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Move_Now = True
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'   if button=  then
If (Button = vbLeftButton) Then
  If Move_Now = True Then
          IMAGE1(Index).Move x + IMAGE1(Index).Left, Y + IMAGE1(Index).Top
'          Label1(Index).Move x + Label1(Index).left, Y + Label1(Index).top
   End If
End If
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Move_Now = False
End Sub


Private Sub mean_command_Click()
    If Frame1.Visible = True Then
            Frame1.Visible = flase
            mean_command.Caption = "&H开关含义说明"
    Else
            Frame1.Visible = True
            mean_command.Caption = "&H掩蔽开关含义"
    End If
End Sub

Private Sub Timer1_Timer()
  Dim x As Integer
   x = read_elec_output(Elec_Output)
  For i = 0 To Elec_Output.Length
        If IMAGE1(i).Status <> Elec_Output.data(i) Then
                IMAGE1(i).Status = Elec_Output.data(i)
        End If
 Next i
End Sub




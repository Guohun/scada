VERSION 5.00
Begin VB.Form elec_disp 
   Caption         =   "Form1"
   ClientHeight    =   5916
   ClientLeft      =   228
   ClientTop       =   756
   ClientWidth     =   9360
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
   ScaleHeight     =   5916
   ScaleWidth      =   9360
   Begin VB.CommandButton Exit_command 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6120
      Top             =   3600
   End
   Begin VB.PictureBox Image1 
      Height          =   735
      Index           =   0
      Left            =   1320
      ScaleHeight     =   684
      ScaleWidth      =   1044
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "elec_disp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public data As Database
Public rec As Recordset
Dim tim 'As idoit 'Integer
Dim total_set  As Integer
Dim Note_Name(300)  As Integer






Private Sub exit_command_Click()
    Unload Me
End Sub


Private Sub Form_Load()
On Error Resume Next
 Width = MAIN_MDI.Width  ' Set width of form.
 Height = MAIN_MDI.Height - 600 ' Set height of form.
 Left = 0
 Top = 0
Call Add_Win(Hwnd)

    Dim i As Integer
    Set data = OpenDatabase("c:\ljxt\ljxt.mdb")
    Set rec = data.OpenRecordset("select * from  elec_disp a ,node_table   b where b.alias=a.name and b.in_out=0 ")
    rec.MoveLast
    total_set = rec.RecordCount
    For i = 1 To total_set
            Load IMAGE1(i)
    Next i
    rec.MoveFirst
    Elec_In
    If china_english = CHINA Then
       exit_command.Caption = "&E退出"
       Caption = "电信号显示"
    Else
       exit_command.Caption = "&Exit"
       Caption = "elect_disp"
    End If
End Sub



Public Sub Elec_In()
Dim i    As Integer
    i = 0
    rec.MoveFirst
    Do While Not rec.EOF
        name0 = rec.Fields("name").Value
        row0 = rec.Fields("row").Value
        col0 = rec.Fields("col").Value
        eof0 = rec.Fields("eof").Value
        lenght0 = rec.Fields("lenght").Value
        type0 = rec.Fields("type").Value
        wei0 = rec.Fields("wei").Value
        china_str0 = rec.Fields("china_str").Value
        china_dat0 = rec.Fields("china_dat").Value
        eng_str0 = rec.Fields("eng_str").Value
        eng_dat0 = rec.Fields("eng_dat").Value
        Note_Name(i) = rec.Fields("node_name")          '所有接点名
        
            If china_english = CHINA Then
                IMAGE1(i).Caption = china_str0
            Else
                IMAGE1(i).Caption = eng_str0
            End If
            IMAGE1(i).Visible = True
            IMAGE1(i).Status = 0
            IMAGE1(i).Left = col0
            IMAGE1(i).Top = row0
             i = i + 1
        rec.MoveNext
    Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    For i = 1 To total_set
            Unload IMAGE1(i)
    Next i
    rec.Close
    data.Close
    Call Del_Win(Hwnd)

End Sub




Private Sub Timer1_Timer()
 Dim i As Integer
   Dim x As Long
     'For i = 0 To gcount - 1
  x = read_elec_input(Elec_Input)
 For i = 0 To Elec_Input.Length
'     MsgBox Elec_Input.data(i)
'    MsgBox Elec_Input.Note_Name(i)
    If Elec_Input.data(i) = 1 Then
         IMAGE1(i).Status = 1
    Else
         IMAGE1(i).Status = 0
     End If
 Next i
End Sub



VERSION 5.00
Begin VB.Form elec_place_Form 
   Caption         =   "电信号录入："
   ClientHeight    =   2772
   ClientLeft      =   1140
   ClientTop       =   1512
   ClientWidth     =   6372
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2772
   ScaleWidth      =   6372
   Begin VB.CommandButton Del_Command 
      Caption         =   "Del"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton UPDATE_Command 
      Caption         =   "UPDATE"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cancel_command 
      Caption         =   "cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton OK_command 
      Caption         =   "&OK"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Disp_Frame 
      Caption         =   "请输入："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      Begin VB.TextBox elec_name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3360
         TabIndex        =   10
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox eng_elec_name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3360
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox node_name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox china_elec_name 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "接点描述名："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "英文电信号名："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "说明：接点名必须为整数"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "接点名："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "中文电信号名："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "elec_place_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const LIGHT_DOING = 0
Const CHECK_DOING = 1
Const NOTHING_DOING = 2

Private Sub cancel_command_Click()
     elec_node_form.Current_china_str = ""
     elec_node_form.Current_node_name = 0
     Edit_Light_Flag = False
     Edit_Check_Flag = False

     Unload Me
End Sub


Private Sub del_command_Click()
        elec_node_form.change_flag = CHANGE
        Call elec_node_form.Del_Elec
        Call elec_node_form.Add_list        'also  edit list
         Edit_Light_Flag = False
         Edit_Check_Flag = False
        Unload Me

End Sub

Private Sub Form_Load()
        node_name.Text = CStr(elec_node_form.Current_node_name)
        china_elec_name.Text = elec_node_form.Current_china_str
        eng_elec_name.Text = elec_node_form.Current_eng_str
        elec_name.Text = elec_node_form.Current_name
        Call Add_Win(Hwnd)

        If elec_node_form.Current_node_name = -1 Then
            node_name.Text = "0"
            Ok_Command.Visible = True
            UPDATE_Command.Visible = False
            del_command.Visible = False
      Else
            Ok_Command.Visible = flase
            UPDATE_Command.Visible = True
            del_command.Visible = True
            
        End If
End Sub




Private Sub Form_Paint()
            If china_english = CHINA Then
                    Disp_Frame.Caption = "请输入 "
                    Ok_Command.Caption = "&O确认"
                    UPDATE_Command.Caption = "&U修改"
                    del_command.Caption = "&D删除"
            Else
                    Disp_Frame.Caption = "Please INput "
                    Ok_Command.Caption = "&Ok"
                    UPDATE_Command.Caption = "&Update"
                    del_command.Caption = "&Del"
            End If
End Sub


Private Sub Form_Terminate()
            Call Del_Win(Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call Del_Win(Hwnd)
End Sub

Private Sub Ok_Command_Click()
        If node_name.Text = "" Then
                MsgBox "必须输入接点名"
                Exit Sub
        End If
        If china_elec_name.Text = "" Then
                MsgBox "必须输入中文名"
                Exit Sub
        End If
        If eng_elec_name.Text = "" Then
                MsgBox "必须输入中文名"
                Exit Sub
        End If
        If elec_name.Text = "" Then
                MsgBox "必须输入接点描述名"
                Exit Sub
        End If
        elec_node_form.change_flag = CHANGE
        elec_node_form.Current_node_name = CInt(node_name.Text)
        elec_node_form.Current_china_str = china_elec_name.Text
        elec_node_form.Current_eng_str = eng_elec_name.Text
        elec_node_form.Current_name = elec_name.Text
        If elec_node_form.Place_Flag = LIGHT_DOING Then
             Call elec_node_form.Add_Light(elec_node_form.Current_row, elec_node_form.Current_col)
        End If
        If elec_node_form.Place_Flag = CHECK_DOING Then
             Call elec_node_form.Add_Check(elec_node_form.Current_row, elec_node_form.Current_col)
        End If
        Call elec_node_form.Add_list
        Unload Me
   End Sub


Private Sub UPDATE_Command_Click()

        If node_name.Text = "" Then
                Call Speak_Error("必须输入接点名")
                Exit Sub
        End If
        If china_elec_name.Text = "" Then
                Call Speak_Error("必须输入中文名")
                Exit Sub
        End If
        If eng_elec_name.Text = "" Then
                Call Speak_Error("必须输入中文名")
                Exit Sub
        End If
        If elec_name.Text = "" Then
                Call Speak_Error("必须输入接点描述名")
                Exit Sub
        End If
        
          Dim i   As Integer
          Dim str   As String
          
      str = node_name.Text
      If Len(str) > 6 Then
                Call Speak_Error("数值太大")
                Exit Sub
      End If
    For i = 1 To Len(str)
         If Asc(Mid(str, i, 1)) < Asc("0") Or Asc(Mid(str, i, 1)) > Asc("9") Then
                Call Speak_Error("必须输入数值")
                Exit Sub
         End If
  Next i
        elec_node_form.change_flag = CHANGE
        elec_node_form.Current_node_name = CInt(node_name.Text)
        elec_node_form.Current_china_str = china_elec_name.Text
        elec_node_form.Current_eng_str = eng_elec_name.Text
        elec_node_form.Current_name = elec_name.Text
        If elec_node_form.Edit_Light = LIGHT_DOING Then
             Call elec_node_form.Edit_Light
        End If
        If elec_node_form.Edit_Flag = check Then
             Call elec_node_form.Edit_Check
        End If
        Call elec_node_form.Add_list        'also  edit list
         Edit_Light_Flag = False
         Edit_Check_Flag = False

        Unload Me

End Sub



VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Prep_Cai_Liao 
   Caption         =   "材料选择"
   ClientHeight    =   4152
   ClientLeft      =   1284
   ClientTop       =   1356
   ClientWidth     =   7356
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4152
   ScaleWidth      =   7356
   Begin Threed.SSCommand Command 
      Height          =   852
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2350
      _ExtentY        =   1503
      _StockProps     =   78
      Caption         =   "SSCommand1"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Picture         =   "Prep_Cai_Liao.frx":0000
   End
   Begin VB.ListBox Cai_LIst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2544
      Left            =   3840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1812
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   5880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2346
      _ExtentY        =   1288
      _StockProps     =   78
      Caption         =   "SSCommand1"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Picture         =   "Prep_Cai_Liao.frx":0452
   End
   Begin Threed.SSCommand Command 
      Height          =   852
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2350
      _ExtentY        =   1503
      _StockProps     =   78
      Caption         =   "SSCommand1"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Picture         =   "Prep_Cai_Liao.frx":08A4
   End
   Begin Threed.SSCommand Command 
      Height          =   852
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2350
      _ExtentY        =   1503
      _StockProps     =   78
      Caption         =   "SSCommand1"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Picture         =   "Prep_Cai_Liao.frx":0CF6
   End
   Begin VB.Line Line4 
      X1              =   1800
      X2              =   1080
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   1080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   1080
      Y1              =   1560
      Y2              =   3240
   End
   Begin VB.Label List_Name 
      Alignment       =   2  'Center
      Caption         =   "list_name"
      Height          =   372
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1812
   End
End
Attribute VB_Name = "Prep_Cai_Liao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mathine_code As Integer
Public wna_code As Integer
Public wna_name As String
Public wna_max   As Single
Dim max_array(0 To 20)      As Single
Dim max_ban(0 To 20)      As Integer


'****************************
'所选秤的重量处理
'****************************
Private Function Add_Ban(ByVal temp As Integer) As Boolean
    Dim db As Database, rs As Recordset
    Dim i%
    Cai_LIst.Clear
    i% = 0
    Add_Ban = False

    Set db = OpenDatabase(Data_Path & "\comm.mdb")
    Set rs = db.OpenRecordset("select code,name,max from send_table where code <=4  and  code>0  and  mathine=" & temp, dbOpenDynaset)
        Do While Not rs.EOF
                Cai_LIst.AddItem vFieldVal(rs.Fields("name").Value)
                max_array(i%) = CSng(vFieldVal(rs.Fields("max").Value))
                max_ban(i%) = rs.Fields("code").Value
                rs.MoveNext
                i% = i% + 1
                Add_Ban = True
        Loop
    rs.Close
    db.Close
End Function




Private Sub cai_list_DblClick()
            If Cai_LIst.ListIndex < 0 Then
                Exit Sub
            End If

        wna_code = max_ban(Cai_LIst.ListIndex) 'CInt(Right(Trim(cai_list.List(cai_list.ListIndex)), 1))
        wna_name = Cai_LIst.List(Cai_LIst.ListIndex)
        wna_max = max_array(index)
        Cai_liao_table_Input.Show
End Sub


Private Sub Cai_LIst_GotFocus()
    If Cai_LIst.ListCount > 0 Then
        Cai_LIst.ListIndex = 1
    End If
End Sub

Private Sub cai_list_KeyPress(KeyAscii As Integer)
            If Cai_LIst.ListIndex < 0 Then
                Exit Sub
            End If
 If KeyAscii = 13 Then
        wna_code = max_ban(Cai_LIst.ListIndex) 'CInt(Right(Trim(cai_list.List(cai_list.ListIndex)), 1))
       
        wna_name = Cai_LIst.List(Cai_LIst.ListIndex)
        wna_max = max_array(index)
        Cai_liao_table_Input.Show
End If
End Sub


Private Sub Cai_LIst_LostFocus()
        Cai_LIst.ListIndex = -1
End Sub

Private Sub Command_Click(index As Integer)
        
        If Add_Ban(index) = False Then Exit Sub
        mathine_code = index
        Cai_LIst.SetFocus
        If China_English = CHINA Then
'            Select Case Index
 '               Case -1:
  '                  list_name.Caption = "秤名："
   '             Case 0:
    '                list_name.Caption = "主机秤名："
     '           Case 1:
      '              list_name.Caption = "机组1秤名："
       '         Case 2:
        '            list_name.Caption = "机组2秤名："
         '   End Select
        Else
          ' Select Case mathine_code
           ' Case -1:
            '    list_name.Caption = "cheng："
            'case 0:
              '  list_name.Caption = "main："
            'Case 1:
             '   list_name.Caption = "mathine1："
           ' Case 2:
            '    list_name.Caption = "mathine2："
        'End Select
       End If
        
End Sub


Private Sub Exit_Command_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
            Me.WindowState = 2
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

'
Private Sub Form_Load()
            mathine_code = -1
            Width = MAIN_MDI.Width  ' Set width of form.
            Height = MAIN_MDI.Height * 3 / 4 ' Set height of form.
            Left = 0
            Top = 0

            Call Add_Win(Hwnd)
End Sub



Private Sub Form_Paint()
    If China_English = CHINA Then
        Caption = "材料输入"
        Exit_Command.Caption = "ESC退出"
        Command(0).Caption = "&0主机"
        Command(1).Caption = "&1机组"
        Command(2).Caption = "&2机组"
        list_name.Caption = "秤名："
    Else
        Caption = "cai_liao_input"
        Exit_Command.Caption = "&Exit"
        Command(0).Caption = "&0Mathin"
        Command(1).Caption = "&1"
        Command(2).Caption = "&2"
        list_name.Caption = "wman name："
    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Call Del_Win(Hwnd)
End Sub



VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form F_Error 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "错误提示窗口"
   ClientHeight    =   3696
   ClientLeft      =   6564
   ClientTop       =   3924
   ClientWidth     =   2688
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   15.4
   ScaleMode       =   4  'Character
   ScaleWidth      =   22.4
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\Ljxt.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "error_num"
      Top             =   2880
      Width           =   2052
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2160
      ItemData        =   "f_error.frx":0000
      Left            =   0
      List            =   "f_error.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   2652
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2412
      _Version        =   65536
      _ExtentX        =   4254
      _ExtentY        =   656
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin Threed.SSCommand SSCommand2 
         Height          =   372
         Left            =   1320
         TabIndex        =   4
         Top             =   0
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "隐藏窗口"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   372
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "取消报警"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
   End
   Begin VB.OLE OLE1 
      Class           =   "Soundo'LE"
      DataField       =   "china_speaker"
      DataSource      =   "Data1"
      Height          =   492
      Left            =   120
      OleObjectBlob   =   "f_error.frx":0004
      TabIndex        =   2
      Top             =   2880
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "按[E]键切换窗口"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   3372
   End
End
Attribute VB_Name = "F_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'显示报警信息及声音
Public Function bak_Speak_Error()
Static Error_Rec  As Variant
Static Is_in_db As Boolean
Dim Found As Boolean
        
If (Cur_errors.length > 0) Then         '有错误
    Found = False
    If Cur_speaking_error_code <> Cur_errors.Code(0) Then       '新错误
        Speaking_error = True
        SSCommand1.Caption = "取消报警"
        
        '打开错误报警窗口
        F_Error.Width = 2784
        F_Error.Height = 3216
        F_Error.Left = MAIN_MDI.ScaleWidth - F_Error.Width
        F_Error.Top = MAIN_MDI.ScaleHeight - F_Error.Height
        If Not F_Error.Visible Then
             F_Error.Show vbModal
        End If
        F_Error.ZOrder (0)
        
        F_Error.Data1.Recordset.Index = "error_idx"
        F_Error.Data1.Recordset.Seek "=", CStr(Cur_errors.Code(0))
        If Not F_Error.Data1.Recordset.NoMatch Then     'Found error message
            Found = True
            Is_in_db = True
            Error_Rec = F_Error.Data1.Recordset.Bookmark
            F_Error.List1.AddItem " [" + Left(Trim(CStr(Cur_errors.Code(0))), 4) + "] " + Trim(F_Error.Data1.Recordset!china_prompt)
            Tmp = DoEvents()
        Else
            Found = False
            Is_in_db = False
            Error_Rec = -1
            F_Error.List1.AddItem " [" + Left(Trim(CStr(Cur_errors.Code(0))), 4) + "] " + Trim(Cur_errors.prompt(0))
        End If
        F_Error.List1.ListIndex = F_Error.List1.ListCount - 1
        F_Error.List1.ItemData(F_Error.List1.ListIndex) = CStr(Cur_errors.Code(0))
        'kill_error (Cur_errors.Code(0))             'kill error
        
        Cur_speaking_error_code = Cur_errors.Code(0)
    Else                '旧错误
        If Is_in_db Then
            Found = True
        End If
    End If
Else                    ' 无错误
    '取消正在报警的错误
    'If (Cur_speaking_error_code <> -1) And Speaking_error Then
        Found = False
        Cur_speaking_error_code = -1
        'Error_rec = -1
    'End If
End If

'播放声音
If Found Then           '在库中找到错误信息
    If Speaking_error And (Not OLE1.AppIsRunning) Then
        Tmp = DoEvents()
        F_Error.Data1.Recordset.Bookmark = Error_Rec
        On Error Resume Next
        F_Error.OLE1.DoVerb (0)     'play voice
        OLE1.AppIsRunning = False
    End If
'    While Speaking_error And Found
'       f_error.Data1.Recordset.Bookmark = Error_rec
'       tmp = DoEvents()
'       If OLE1.AppIsRunning = False Then
'            On Error Resume Next
'            f_error.OLE1.DoVerb (0)     'play voice
'            OLE1.AppIsRunning = False
'        End If
'    Wend
End If

End Function

'显示报警信息及声音
Public Sub Old_Speak_error()
Dim i As Integer
Dim First_error  As Variant
On Error Resume Next
    If (F_Error.Visible = False) Then
        F_Error.Show
    End If
    F_Error.ZOrder (0)
    
    'speaking voice
    F_Error.Data1.Recordset.Index = "error_idx"
    
        For T = 0 To Cur_errors.length - 1
            Found = False
            For n = 0 To Last_errors.length - 1
                If Cur_errors.Code(T) <> Last_errors.Code(n) Then
                    Found = True
                    Exit For
                End If
            Next n
            If Not Found Then       'new error
                Cur_speaking_error_code = Cur_errors.Code(T)
                Last_errors.Code(length) = Cur_errors.Code(T)
                Last_errors.Code(length) = Cur_errors.prompt(T)
                Last_errors.length = Last_errors.length + 1
            End If
        Next T
    '查找第一个错误
    If (Cur_errors.length > 0) Then
        If Cur_speaking_error_code <> Cur_errors.Code(0) Then
            F_Error.Data1.Recordset.Seek "=", CStr(Cur_errors.Code(0))
            If Not F_Error.Data1.Recordset.NoMatch Then     'found error message
                First_error = F_Error.Data1.Recordset.Bookmark
                F_Error.List1.AddItem " [" + Left(Trim(CStr(Cur_errors.Code(0))), 4) + "] " + Trim(F_Error.Data1.Recordset!china_prompt)
            Else
                First_error = -1
                F_Error.List1.AddItem " [" + Left(Trim(CStr(Cur_errors.Code(0))), 4) + "] " + Trim(Cur_errors.prompt(0))
            End If
            F_Error.List1.ListIndex = F_Error.List1.ListCount - 1
            F_Error.List1.ItemData(F_Error.List1.ListIndex) = CStr(Cur_errors.Code(0))
            'kill_error (Cur_errors.Code(0))             'kill error
            
            Cur_speaking_error_code = Cur_errors.Code(0)
        End If
    
        '查找剩余的错误
        For i = 1 To Cur_errors.length - 1
            F_Error.Data1.Recordset.Seek "=", CStr(Cur_errors.Code(i))
            If Not F_Error.Data1.Recordset.NoMatch Then     'error found
                F_Error.List1.AddItem " [" + Left(Trim(CStr(Cur_errors.Code(i))), 4) + "] " + Trim(F_Error.Data1.Recordset!china_prompt)
            Else
                F_Error.List1.AddItem " [" + Left(Trim(CStr(Cur_errors.Code(i))), 4) + "] " + Trim(Cur_errors.prompt(i))
            End If
            F_Error.List1.ItemData(F_Error.List1.ListCount - 1) = CStr(Cur_errors.Code(i))
            'kill_error (Cur_errors.Code(i))
        Next i
    Else        'no found any error
        Found = False
    End If
    
    '播放声音
    If First_error <> -1 Then
        F_Error.Data1.Recordset.Bookmark = First_error
        'tmp = DoEvents()
        If Speaking_error And (Not OLE1.AppIsRunning) Then
            On Error Resume Next
            F_Error.OLE1.DoVerb (0)     'play voice
            OLE1.AppIsRunning = False
        End If
    End If
End Sub




Private Sub Form_Deactivate()
'Me.ZOrder 0
End Sub

Private Sub Form_Initialize()
'Cur_speaking_error_code = Cur_errors.Code(0)
End Sub

Private Sub Form_Load()
Speaking_error = True
        
F_Error.Width = 2784
F_Error.Height = 3216
F_Error.Left = MAIN_MDI.ScaleWidth - F_Error.Width
F_Error.Top = MAIN_MDI.ScaleHeight - F_Error.Height
End Sub


Private Sub Form_Resize()
If Me.WindowState = 0 Then
    F_Error.Width = 2784
    F_Error.Height = 3216
Else
    Me.WindowState = 0
End If
End Sub

Private Sub list1_DblClick()
Call list1_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Here As Variant
Select Case KeyCode
    Case vbKeyE
        F_ShowLc.SetFocus
    Case vbKeyEscape
        F_Error.Hide
    Case vbKeyReturn
'        Here = f_error.Data1.Recordset.Bookmark
        If Not F_Error_Help.Visible Then
            F_Error.Data1.Recordset.Index = "error_idx"
            F_Error.Data1.Recordset.Seek "=", List1.ItemData(List1.ListIndex)
            If Not F_Error.Data1.Recordset.NoMatch Then
                F_Error_Help.Label2.Caption = List1.List(List1.ListIndex)
                F_Error_Help.Text1.Text = F_Error.Data1.Recordset!china_help
                F_Error_Help.Show vbModal
            Else
                F_Error_Help.Label2.Caption = List1.List(List1.ListIndex)
                F_Error_Help.Text1.Text = ""
            End If
        End If
'        f_error.Data1.Recordset.Bookmark = Here
End Select
End Sub





Private Sub SSCommand1_Click()
    If Speaking_error Then
        Speaking_error = False
        SSCommand1.Caption = "启动报警"
    Else
        Speaking_error = True
        SSCommand1.Caption = "取消报警"
    End If
End Sub


Private Sub SSCommand2_Click()
    F_Error.Hide
End Sub



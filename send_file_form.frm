VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form send_file_form 
   ClientHeight    =   3636
   ClientLeft      =   2016
   ClientTop       =   2040
   ClientWidth     =   5772
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3636
   ScaleWidth      =   5772
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Get_Command 
      Caption         =   "Get_Command"
      Height          =   972
      Left            =   3840
      TabIndex        =   1
      Top             =   1440
      Width           =   1812
   End
   Begin VB.CommandButton send_command 
      Caption         =   "Command1"
      Height          =   972
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   1812
   End
   Begin MSComDlg.CommonDialog openlog 
      Left            =   2856
      Top             =   1704
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   492
      Left            =   2520
      TabIndex        =   2
      Top             =   2880
      Width           =   852
      _Version        =   65536
      _ExtentX        =   1503
      _ExtentY        =   868
      _StockProps     =   78
      Caption         =   "ESC"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.Label Prompt_Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   1872
      TabIndex        =   3
      Top             =   480
      Width           =   2220
   End
End
Attribute VB_Name = "send_file_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancelsend     As Boolean
   

Private Sub Exit_Command_Click()
    cancelsend = True
    Prompt_Label.Caption = "正常"
    Call Sleep(100)
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
        Call Add_Win(Hwnd)
End Sub

Private Sub Form_Paint()

        If China_English = CHINA Then
                Prompt_Label.Caption = "正常"
                Caption = "发送接收文件"
                Send_Command.Caption = "&P发送数据文件"
                Get_Command.Caption = "&C接收数据文件"
                Exit_Command.Caption = "ESC退出"
        Else
                Prompt_Label.Caption = "Normal"
                Exit_Command.Caption = "ESC Exit "
                Send_Command.Caption = "&PSend data file"
                Get_Command.Caption = "&C   recevier"
        End If
End Sub

Private Sub Form_Terminate()
Call Del_Win(Hwnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call Del_Win(Hwnd)
        Call sio_close(Send_File_Port)
        Call sio_close(Get_File_Port)
End Sub

Private Sub Get_Command_Click()
   'On Error Resume Next
   Dim hSend, BSize, LF&
   Dim temp$
   Dim Err$
   Dim buffer As String * 1
   
   Send_Command.Enabled = False
   
   ' Get the text filename from the user.
     openlog.DialogTitle = "Send Text File"
     openlog.Filter = "Text Files (*.TXT)|*.txt|All Files (*.*)|*.*"
      openlog.FileName = ""
      openlog.ShowOpen
      temp$ = openlog.FileName
      ' If the file doesn't exist, go back.
      If temp$ = "" Then
               MsgBox temp$ + "非法文件!", 48
               Exit Sub
      End If
      ret = Len(Dir$(temp$))
      If ret Then
                Call FileCopy(temp$, temp$ & ".bak")
                Call Kill(temp$)
      Else
         MsgBox temp$ + " 新文件!", 48
      End If

   ' Open the log file.
   hSend = FreeFile
    'Open "TESTFILE" For Output Shared As #1
    Open temp$ For Binary Access Write As hSend
        If (sio_open(Get_File_Port) < 0) Then
                Call Speak_Error("打开口" & Get_File_Port & "错误")
        End If
        
        If (sio_ioctl(Get_File_Port, B9600, BIT_8) < 0) Then
            Call Speak_Error("口" & Get_File_Port & "设置错误")
        End If

     cancelsend = False
      LF& = 60
      temp$ = space$(20)
      Do While Not cancelsend
         Do While (sio_iqueue(Get_File_Port) <= 0) And Not cancelsend
            Prompt_Label.Caption = "...等待..."
            DoEvents
            Prompt_Label.Caption = "...........等待............"
         Loop
         LF& = sio_read(Get_File_Port, buffer, 1)
         If LF& < 1 Then
               Exit Do
         End If
         'put #hSend, temp$
         Put #hSend, , buffer
          DoEvents
    Loop
   Close hSend
   
   Call sio_close(Get_File_Port)
   Send_Command.Enabled = True
   cancelsend = True
End Sub


Private Sub Send_Command_Click()
   'On Error Resume Next
   Dim hSend, BSize, LF&
   Dim Err$
   Dim Send_Size%
   Dim buffer   As String * 1
   
   Send_Command.Enabled = False
   openlog.DialogTitle = "Send Text File"
   openlog.Filter = "Text Files (*.TXT)|*.txt|All Files (*.*)|*.*"
   Do
      openlog.FileName = ""
      openlog.ShowOpen
      'If Err = Cancel Then Exit Sub
      temp$ = openlog.FileName
      ret = Len(Dir$(temp$))
'      If Err Then
 '        MsgBox Err, 48
  '       Send_Command.Enabled = True
   '      Exit Sub
    '  End If
      If temp$ = "" Then
               MsgBox temp$ + "非法文件!", 48
               Exit Sub
      End If
    
      If ret Then
         Exit Do
      Else
         MsgBox temp$ + " not found!", 48
      End If
   Loop

   
   hSend = FreeFile
   Open temp$ For Binary Access Read As hSend
        If (sio_open(Send_File_Port) < 0) Then
                Call Speak_Error("打开口" & Send_File_Port & "错误")
        End If
        
        If (sio_ioctl(Send_File_Port, B9600, BIT_8) < 0) Then
            Call Speak_Error("口" & Send_File_Port & "设置错误")
        End If
      cancelsend = False
      
      Do While Not EOF(hSend) And Not cancelsend
          
          Get hSend, , buffer
          Send_Size = sio_write(Send_File_Port, buffer, 1)
          Do While cancelsend = False And Send_Size > 0
                Prompt_Label.Caption = "...等待..."
                DoEvents
                Send_Size = sio_oqueue(Send_File_Port)
                Prompt_Label.Caption = ".......等待......."
          Loop
          DoEvents
      Loop
   Close hSend
   Send_Command.Enabled = True
   cancelsend = True
   Call sio_close(Send_File_Port)
   Call Speak_Error("传输完毕")
End Sub


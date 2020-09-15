VERSION 5.00
Begin VB.Form song 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Song"
   ClientHeight    =   3000
   ClientLeft      =   1536
   ClientTop       =   1632
   ClientWidth     =   6684
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3000
   ScaleWidth      =   6684
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdRecStop 
      Caption         =   "&Record Stop"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdRecStart 
      Caption         =   "&Record  Start"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Current State: Ready......"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
   End
End
Attribute VB_Name = "song"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WaveFile As String
Dim flgRecording As Integer
Dim flgPlaying As Integer
Dim flgPause As Integer
Private Sub Command1_Click()

End Sub


'* * 程序结束 * *
Private Sub cmdExit_Click()
    If flgRecording Then cmdRecStop_Click
    If flgPlaying Then cmdStop_Click
    Unload Me
End Sub


'* * 暂停播放* *
Private Sub cmdPause_Click()
  If flgPause Then
   '重新播放
   i = mciExecute("play sound")
   cmdPause.Caption = "pause"
   flgPause = False
   Label1 = "Current State: Playing..."
  Else
  '暂停播放
  i = mciExecute("pause sound")
   cmdPause.Caption = "Continue"
   flgPause = True
   Label1 = "Current State: Pause..."
 End If
End Sub


'* *播放语音文件  * *
Private Sub cmdPlay_Click()
     '打开 Waveform 语音设备
     If mciExecute("open+ WaveFile + alias sound") = False Then
         MsgBox "Can not open text.wav"
         Exit Sub
     End If
     '播放语音
     i = mciExecute("play sound") '
     Label = "Current State: Playing..."
     '改变按扭状态
     cmdRecStart.Enabled = False
     cmdRecStop.Enabled = False
     cmdPlay.Enabled = False
     cmdPause.Enabled = True
     cmdStop.Enabled = True
     flgPlaying = True
End Sub


'* * 录音 * *
Private Sub cmdRecStart_Click()
   '检测暂存文件TEST.WAV是否存在
   If Dir$(WaveFile) <> "" Then
  '如果TEST.WAV 在 , 将之删除
      Kill WaveFile
  End If
   '建立一个新文件
   If mciExecute("open new type waveaudio alias sound") = False Then
         MsgBox "Can not creat new file"
         Exit Sub
     End If
     i = mciExecute("record sound")
     Label1 = "Current State: Playing..."
     '改变按扭状态
     cmdRecStart.Enabled = False
     cmdRecStop.Enabled = True
     cmdPlay.Enabled = False
     cmdPause.Enabled = False
     cmdStop.Enabled = False
     
     cmdRecStop.SetFocus
     flgRecording = True
End Sub




'* * 停止录音* *
Private Sub cmdRecStop_Click()
     i = mciExecute("stop sound")
     i = mciExecute("save sound" + WaveFile)
     i = mciExecute("close sound")
     Label1 = "Current State: Ready..."
     '改变按扭状态
     cmdRecStart.Enabled = True
     cmdRecStop.Enabled = False
     cmdPlay.Enabled = True
     cmdPause.Enabled = False
     cmdStop.Enabled = False
     
     cmdPlay.SetFocus
     flgRecording = False
     
End Sub


' * * 停止语音播放 * *
Private Sub cmdStop_Click()
      i = mciExecute("stop sound")
      i = mciExecute("close sound")
     Label1 = "Current State: Ready..."
     '改变按扭状态
     cmdRecStart.Enabled = True
     cmdRecStop.Enabled = False
     cmdPlay.Enabled = True
     cmdPause.Enabled = False
     cmdStop.Enabled = False
     
     cmdPlay.SetFocus
     flgPlaying = False
End Sub


Private Sub Form_Load()
   WaveFile = App.Path + "\TEST.WAVE"
   flgRecording = False
   flgPlaying = False
   flhPause = False
   cmdRecStop.Enabled = False
     cmdPause.Enabled = False
     cmdStop.Enabled = False
   '检测TEST.WAV是否存在
   If Dir$(WaveFile) = "" Then
      cmdPlay.Enabled = False
   End If
   Label1 = "Current State: Ready ...."
End Sub



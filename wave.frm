VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmWave 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3204
   ClientLeft      =   1356
   ClientTop       =   3984
   ClientWidth     =   6588
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3204
   ScaleWidth      =   6588
   Begin Threed.SSFrame SSFrame1 
      Height          =   1572
      Left            =   4560
      TabIndex        =   7
      Top             =   960
      Width           =   1812
      _Version        =   65536
      _ExtentX        =   3196
      _ExtentY        =   2773
      _StockProps     =   14
      Caption         =   "录音选择"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
      Begin Threed.SSOption Check1 
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   78
         Caption         =   "SSOption1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.59
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSOption Check1 
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   78
         Caption         =   "SSOption1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.59
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
   End
   Begin VB.HScrollBar hsbWave 
      Height          =   255
      Left            =   360
      Max             =   100
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin MCI.MMControl mciWave 
      Height          =   492
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   868
      _Version        =   327680
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin Threed.SSCommand Exit_Command 
      Cancel          =   -1  'True
      Height          =   732
      Left            =   2760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
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
      Font3D          =   1
      Picture         =   "wave.frx":0000
   End
   Begin Threed.SSCommand Update_Command 
      Height          =   732
      Left            =   960
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "NoEdit"
      Top             =   1440
      Width           =   1332
      _Version        =   65536
      _ExtentX        =   2346
      _ExtentY        =   1291
      _StockProps     =   78
      Caption         =   "Save"
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
      Font3D          =   1
      Picture         =   "wave.frx":0452
   End
   Begin VB.Label NowLabel 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   852
   End
   Begin VB.Label lblWaveSec 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   972
   End
   Begin VB.Label lblWave 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const conInterval = 20
Const conIntervalPlus = 55
Dim change_flag    As Boolean
Dim CurrentValue As Double

'打开退出键
Private Sub AI_EXIT_Click()
    Unload frmWave
End Sub

'打开录音机
Private Sub AI_OPEN(ByVal FileName As String)
    Dim msec As Double
    ' Set the number of milliseconds between successive
    ' StatusUpdate events.
    mciWave.UpdateInterval = 0

    ' Display the File Open dialog box.
    ' If the device is open, close it.
    If Not mciWave.mode = vbMCIModeNotOpen Then
        mciWave.Command = "Close"
    End If

    ' Open the device with the new filename.
       
    
    On Error GoTo MCI_ERROR
   ' 检测暂存文件是否存在
   If Dir$(FileName) <> "" Then
          mciWave.FileName = FileName
          mciWave.Command = "Open"
    Else
         '建立一个新文件
          mciWave.FileName = Data_Path & "\wave\default.wav"
          FileCopy Data_Path & "\wave\default.wav", FileName
          mciWave.Command = "Open"
  End If



    On Error GoTo 0
    
    
    ' Set the timing labels on the form.
    mciWave.TimeFormat = vbMCIFormatMilliseconds
    lblWave.Caption = "0.0"
    msec = (CDbl(mciWave.Length) / 1000)
    lblWaveSec.Caption = Format$(msec, "0.00")

    ' Set the scrollbar values.
    hsbWave.Value = 0
    CurrentValue = 0#
    Exit Sub

MCI_ERROR:
    DispWaveErrorMessageBox
    Resume MCI_EXIT

MCI_EXIT:
    Unload frmWave
    
End Sub
'录音停止
Private Sub cmdRecStop_Click()
    mciWave.UpdateInterval = 0
    ' Reset the scrollbar values.
    hsbWave.Value = 0
    CurrentValue = 0#
    mciWave.Command = "stop"
End Sub

Private Sub Check1_Click(index As Integer, Value As Integer)
        mciWave.RecordMode = index
End Sub

Private Sub Exit_Command_Click()
        Unload Me
End Sub

Private Sub Form_Activate()
    If mciWave.Tag = "" Then
            AI_OPEN (Me.Tag)
            mciWave.Tag = "123"
            change_flag = UNCHANGE
            Check1(0).SetFocus
            
    End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF5
                            Call Update_Command_Click
                Case vbKeyEscape
                            Unload Me
        End Select
        
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()

    ' Force the multimedia MCI control to complete before returning
    ' to the application.
    
    frmWave.mciWave.Wait = True
End Sub

Private Sub Form_Paint()
    If China_English = CHINA Then
            'Caption = "录放音处理"
    Else
            'Caption = "record and play wave"
    End If
        If China_English = CHINA Then
                Exit_Command.Caption = "ESC退出"
                Update_Command.Caption = "F5保存"
                Check1(0).Caption = "&I插入"
                Check1(1).Caption = "&O覆盖"
        Else
                Check1(0).Caption = "&Insert"
                Check1(1).Caption = "&OverWrite"
                Exit_Command.Caption = "ESC "
                Update_Command.Caption = "F5 "
            
        End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
            If change_flag = CHANGE Then
                    answer = MsgBox(" 声音已修改  确定退出吗？退出时要等待", 48 + vbYesNo, "提示")
                    If answer = vbYes Then
                                Unload Me
                                Cancel = 0
                    Else
                                Cancel = 1
                    End If
            End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub hsbWave_Change()
        If Not mciWave.mode = 526 Then
                mciWave.From = mciWave.Length * hsbWave.Value / (hsbWave.Max - hsbWave.Min)
                CurrentValue = hsbWave.Value * CInt(mciWave.Length / 100) - 1
        End If
End Sub

Private Sub hsbWave_KeyDown(KeyCode As Integer, Shift As Integer)
            Select Case KeyCode
                Case vbKeyF5
                            Call Update_Command_Click
                            key_code = 0
                Case vbKeyEscape
                            Unload Me
                            key_code = 0
        End Select
End Sub

Private Sub mciWave_PauseClick(Cancel As Integer)
    ' Set the number of milliseconds between successive
    ' StatusUpdate events.
    mciWave.UpdateInterval = 0
End Sub

Private Sub mciWave_PlayClick(Cancel As Integer)
    ' Set the number of milliseconds between successive
    ' StatusUpdate events.
    mciWave.UpdateInterval = conInterval
End Sub


Private Sub mciWave_PrevClick(Cancel As Integer)
    ' Set the number of milliseconds between successive
    ' StatusUpdate events.
    mciWave.UpdateInterval = 0
    
    ' Reset the scrollbar values.
    hsbWave.Value = 0
    CurrentValue = 0#
    
    mciWave.Command = "Prev"
End Sub




Private Sub mciWave_RecordClick(Cancel As Integer)
   If mciWave.CanRecord = False Then
          Call Speak_Error("不能录音")
          Cancel = True
          Exit Sub
   End If
        change_flag = CHANGE
        mciWave.UpdateInterval = conInterval
End Sub


Private Sub mciWave_StatusUpdate()
    Dim Value As Long
    Dim msc    As Double
    

    ' If the device is not playing, reset to the beginning.
    'If Not mciWave.mode = 527 Then
        'hsbWave.value = hsbWave.Max
    '    mciWave.UpdateInterval = 0
     '   Exit Sub
   ' End If
    
    If (Not mciWave.mode = 526) And (Not mciWave.mode = 527) Then
        hsbWave.Value = hsbWave.Max
        mciWave.UpdateInterval = 0
        Exit Sub
    End If
    
    ' Determine how much of the file has played.  Set a
    ' value of the scrollbar between 0 and 100.
    
    CurrentValue = CurrentValue + conIntervalPlus
    Value = CInt((CurrentValue / mciWave.Length) * 100)
        
        
    If Value > hsbWave.Max Then
        Value = hsbWave.Max
     '   mciWave.UpdateInterval = 0
    End If
    
    hsbWave.Value = Value
    msc = (CDbl(CurrentValue) / 1000)
    NowLabel.Caption = Format$(msc, "0.00")
    'NowLabel.Caption=
End Sub


Private Sub mciWave_StopClick(Cancel As Integer)
    mciWave.TimeFormat = vbMCIFormatMilliseconds
    lblWave.Caption = "0.0"
    msec = (CDbl(mciWave.Length) / 1000)
    lblWaveSec.Caption = Format$(msec, "0.00")
    mciWave.UpdateInterval = 0
End Sub


Private Sub Update_Command_Click()
        mciWave.UpdateInterval = 0
        mciWave.FileName = frmWave.Tag
        mciWave.Command = "save"
        change_flag = UNCHANGE
End Sub

VERSION 4.00
Begin VB.Form ConfigScrn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Communication Settings"
   ClientHeight    =   3840
   ClientLeft      =   1095
   ClientTop       =   1440
   ClientWidth     =   4980
   Height          =   4245
   Left            =   1035
   ScaleHeight     =   3840
   ScaleWidth      =   4980
   Top             =   1095
   Width           =   5100
   Begin VB.OptionButton ComPort 
      Caption         =   "Com1"
      Height          =   252
      Index           =   1
      Left            =   1920
      TabIndex        =   31
      Top             =   2340
      Width           =   852
   End
   Begin VB.OptionButton ComPort 
      Caption         =   "Com4"
      Height          =   252
      Index           =   4
      Left            =   1920
      TabIndex        =   30
      Top             =   3240
      Width           =   852
   End
   Begin VB.OptionButton ComPort 
      Caption         =   "Com3"
      Height          =   252
      Index           =   3
      Left            =   1920
      TabIndex        =   29
      Top             =   2940
      Width           =   852
   End
   Begin VB.OptionButton ComPort 
      Caption         =   "Com2"
      Height          =   252
      Index           =   2
      Left            =   1920
      TabIndex        =   28
      Top             =   2640
      Width           =   852
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Baud Rate"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3315
      Begin VB.OptionButton Baud3 
         Caption         =   "300"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Baud6 
         Caption         =   "600"
         Height          =   255
         Left            =   1260
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Baud12 
         Caption         =   "1200"
         Height          =   255
         Left            =   2220
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Baud24 
         Caption         =   "2400"
         Height          =   255
         Left            =   300
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Baud48 
         Caption         =   "4800"
         Height          =   255
         Left            =   1260
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Baud96 
         Caption         =   "9600"
         Height          =   255
         Left            =   2220
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3780
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Data Bits"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   1260
      Width           =   1275
      Begin VB.OptionButton Data7 
         Caption         =   "7"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Data8 
         Caption         =   "8"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Stop Bits"
      Height          =   615
      Left            =   1740
      TabIndex        =   12
      Top             =   1260
      Width           =   1335
      Begin VB.OptionButton Stop1 
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   435
      End
      Begin VB.OptionButton Stop2 
         Caption         =   "2"
         Height          =   255
         Left            =   780
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "&Echo"
      Height          =   615
      Left            =   3300
      TabIndex        =   15
      Top             =   1260
      Width           =   1455
      Begin VB.OptionButton EchoOff 
         Caption         =   "Off"
         Height          =   315
         Left            =   780
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton EchoOn 
         Caption         =   "On"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "&Parity"
      Height          =   1575
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   1275
      Begin VB.OptionButton NoParity 
         Caption         =   "None"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton OddParity 
         Caption         =   "Odd"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton EvenParity 
         Caption         =   "Even"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   900
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Com Port"
      Height          =   1575
      Left            =   1740
      TabIndex        =   22
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Caption         =   "&Flow Control"
      Height          =   1575
      Left            =   3300
      TabIndex        =   23
      Top             =   2040
      Width           =   1455
      Begin VB.OptionButton NoFlow 
         Caption         =   "None"
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton XonFlow 
         Caption         =   "Xon/Xoff"
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton RTSFlow 
         Caption         =   "RTS"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   900
         Width           =   735
      End
      Begin VB.OptionButton BothFlow 
         Caption         =   "Xon/RTS"
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   1200
         Width           =   1155
      End
   End
End
Attribute VB_Name = "ConfigScrn"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' Communication Settings Configuration Form

' Copyright (c) 1994, Crescent Software

DefInt A-Z

Dim Shared NewPort                 ' Temporary configuration settings.
Dim Shared NewBaud$, NewParity$
Dim Shared NewData$, NewStop$
Dim Shared NewShake
Dim Changed As Integer
Dim gsDBName As Database
Dim gsRecordComm As Recordset

' 1200 baud option button.
Private Sub Baud12_Click()
    NewBaud$ = "1200"
    Changed = True
End Sub

' 2400 baud option button.
Private Sub Baud24_Click()
    NewBaud$ = "2400"
    Changed = True
End Sub

' 300 baud option button.
Private Sub Baud3_Click()
    NewBaud$ = "300"
    Changed = True
End Sub

' 4800 baud option button.
Private Sub Baud48_Click()
    NewBaud$ = "4800"
    Changed = True
End Sub

' 600 baud option button.
Private Sub Baud6_Click()
    NewBaud$ = "600"
    Changed = True
End Sub

' 9600 baud option button.
Private Sub Baud96_Click()
    NewBaud$ = "9600"
    Changed = True
End Sub

' Both RTS and Xon/Xoff handshaking option button.
Private Sub BothFlow_Click()
    NewShake = 3
    Changed = True
End Sub

' Cancel button actions.
Private Sub CancelButton_Click()
    Unload ConfigScrn
End Sub

Private Sub ComPort_Click(Index As Integer)
    NewPort = Index
    Changed = True
End Sub

' 7 data bits option button.
Private Sub Data7_Click()
    NewData$ = "7"
    Changed = True
End Sub

' 8 data bits option button.
Private Sub Data8_Click()
    NewData$ = "8"
    Changed = True
End Sub

' Echo Off option button.
Private Sub EchoOff_Click()
    Echo = 0
    Changed = True
End Sub

' Echo On option button.
Private Sub EchoOn_Click()
    Echo = True
    Changed = True
End Sub

' Even Parity option button.
Private Sub EvenParity_Click()
    NewParity$ = "E"
    Changed = True
End Sub

' Initialize and display the configuration form.
Private Sub Form_Load()


    ' Get the current port.
    Changed = False
    Port = Form1.MSComm1.CommPort
    ConfigScrn.ComPort(Port).Value = True   ' Set the option button.

    ' Get the current baud.
    FirstComma = InStr(Form1.MSComm1.Settings, ",")
    baud$ = Left(Form1.MSComm1.Settings, FirstComma - 1)
    
    Select Case Val(baud$)                  ' Select baud.
    Case 300                                ' Set the active baud option button.
        ConfigScrn.Baud3.Value = True
    Case 600
        ConfigScrn.Baud6.Value = True
    Case 1200
        ConfigScrn.Baud12.Value = True
    Case 2400
        ConfigScrn.Baud24.Value = True
    Case 4800
        ConfigScrn.Baud48.Value = True
    Case 9600
        ConfigScrn.Baud96.Value = True
    End Select

    ' Get the current parity.
    parity$ = Mid$(Form1.MSComm1.Settings, FirstComma + 1, 1)
    
    Select Case UCase$(parity$)              ' Select parity.
    Case "N"                                 ' Set the active parity option button.
        ConfigScrn.NoParity.Value = True
    Case "E"
        ConfigScrn.EvenParity.Value = True
    Case "O"
        ConfigScrn.OddParity.Value = True
    End Select

    
    ' Get the data bits.
    SecondComma = FirstComma + 2
    DBits$ = Mid$(Form1.MSComm1.Settings, SecondComma + 1, 1)
    Select Case Val(DBits$)                 ' Select data bits.
    Case 7                                  ' Set the active choice option button.
        ConfigScrn.Data7.Value = True
    Case 8
        ConfigScrn.Data8.Value = True
    End Select

    ' Get the stop bits.
    ThirdComma = SecondComma + 2
    SBits$ = Mid$(Form1.MSComm1.Settings, ThirdComma + 1, 1)
    Select Case Val(SBits$)                 ' Select stop bits.
    Case 1                                  ' Set the active choice option button.
        ConfigScrn.Stop1.Value = True
    Case 2
        ConfigScrn.Stop2.Value = True
    End Select

    
    Select Case Form1.MSComm1.Handshaking
    Case 0                                  ' Set active choice option button.
        ConfigScrn.NoFlow.Value = True
    Case 1
        ConfigScrn.XonFlow.Value = True
    Case 2
        ConfigScrn.RTSFlow.Value = True
    Case 3
        ConfigScrn.BothFlow.Value = True
    End Select

    If Echo Then
        ConfigScrn.EchoOn.Value = True
    Else
        ConfigScrn.EchoOff.Value = True
    End If
End Sub

' No handshaking option button.
Private Sub NoFlow_Click()
    NewShake = 0
    Changed = True
End Sub

' No Parity option button.
Private Sub NoParity_Click()
    NewParity$ = "N"
    Changed = True
End Sub

' Odd Parity option button.
Private Sub OddParity_Click()
    NewParity$ = "O"
    Changed = True
End Sub

' OK button actions.
Private Sub OkButton_Click()
    On Error Resume Next
    
    OldPort = Form1.MSComm1.CommPort
    If NewPort <> OldPort Then                    ' If the port number changes, close the old port.
        If Form1.MSComm1.PortOpen Then
           Form1.MSComm1.PortOpen = False
           ReOpen = True
        End If

        Form1.MSComm1.CommPort = NewPort          ' Set the new port number.

        If Err = 0 Then
           If ReOpen Then
              Form1.MSComm1.PortOpen = True
              'Form1.MOpen.Checked = Form1.MSComm1.PortOpen
              'Form1.MSendText.Enabled = Form1.MSComm1.PortOpen
           End If
        End If
        If Err Then
           MsgBox Error$, 48
           Form1.MSComm1.CommPort = OldPort
           Exit Sub
        End If
    End If
    
    Form1.MSComm1.Settings = NewBaud$ + "," + NewParity$ + "," + NewData$ + "," + NewStop$
    If Err Then
       MsgBox Error$, 48
       Exit Sub
    End If
    Form1.MSComm1.Handshaking = NewShake
    If Err Then
       MsgBox Error$, 48
       Exit Sub
    End If
    If Changed Then
      gsRecordComm.Fields("port") = Str(NewPort)
      gsRecordComm.Fields("bound") = NewBaud$
      gsRecordComm.Fields("data") = NewData$
      gsRecordComm.Fields("stop") = NewStop$
      gsRecordComm.Fields("echo") = Echo
      gsRecordComm.Fields("parity") = NewParity$
      gsRecordComm.Fields("control") = NewShake
      gsRecordComm.Update
    End If
    Unload ConfigScrn                               ' Unload the configuration form.

End Sub

' RTS handshaking option button.
Private Sub RTSFlow_Click()
    NewShake = 2
    Changed = True
End Sub

' 1 stop bit option button.
Private Sub Stop1_Click()
    NewStop$ = "1"
    Changed = True
End Sub

' 2 stop bits option button.
Private Sub Stop2_Click()
    NewStop$ = "2"
    Changed = True
End Sub

' XON handshaking option button.
Private Sub XonFlow_Click()
    NewShake = 1
    Changed = True
End Sub


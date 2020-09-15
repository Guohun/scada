VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   1635
   ClientTop       =   1755
   ClientWidth     =   6690
   Height          =   4560
   Left            =   1575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6690
   Top             =   1410
   Width           =   6810
   Begin VB.CommandButton Setting 
      Caption         =   "S&ettings"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6120
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CDTimeout       =   0
      CommPort        =   1
      CTSTimeout      =   0
      DSRTimeout      =   0
      DTREnable       =   -1  'True
      Handshaking     =   0
      InBufferSize    =   1024
      InputLen        =   0
      Interval        =   1000
      NullDiscard     =   0   'False
      OutBufferSize   =   512
      ParityReplace   =   "?"
      RThreshold      =   0
      RTSEnable       =   0   'False
      Settings        =   "9600,n,8,1"
      SThreshold      =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim gsDBName As Database
Dim gsRecordSend As Recordset
Dim gsRecordComm As Recordset
Dim Ret

Private Sub Form_Load()
  FirstDo
  SendData
  ReciveData
End Sub



'Send the message from port that pointed in the table
Public Sub SendData()
  'Select the port to send data
  Dim BSize
  Dim temp As String
  ''Select Case gsRecordSend.field("part").Value
  ''  Case "com1":
  ''    mscomm1.CommPort = 1
  ''  Case "com2":
  ''    mscomm1.CommPort = 2
  ''  Case "com3":
  ''    mscomm1.CommPort = 3
  ''  Case "com4":
  ''    mscomm1.CommPort = 4
  ''End Select
  MSComm1.CommPort = 2
  'setting the bound and other arguements
  'This settings should read the data from comm_table
  MSComm1.Settings = "9600,N,8,1"
  'Tell the control to read entire buffer when Input is used.
  MSComm1.InputLen = 0
  MSComm1.Handshaking = 1

  ' Open the port
  MSComm1.PortOpen = True
  ' Send the data to the com
  BSize = MSComm1.OutBufferCount
  'temp$ = Space$(BSize)
  temp = "Hi, Have you recived my message?" '+ Chr$(13)
  ''temp$ = gsRecordSend.Fields("true_data").Value
  MSComm1.Output = temp
  MsgBox " Message sent!" & temp
  If Err Then
    MsgBox "Error", 48
    Exit Sub
  End If
  'Wait for all data to be sent
  MsgBox MSComm1.OutBufferCount
  Do
    Ret = DoEvents()
    MsgBox MSComm1.OutBufferCount
  Loop Until MSComm1.OutBufferCount = 0
End Sub

Public Sub ReciveData()
' Wait for data to come back to the serial port.
    Dim inString As String
    MsgBox " Waiting for answer!"
    Do
        Ret = DoEvents()
    Loop Until MSComm1.InBufferCount >= 2
    ' Read the "OK" response data in the serial port.
    inString = MSComm1.Input
    MsgBox "Message Recived--" & inString
' Close the serial port.
    MSComm1.PortOpen = False

End Sub

Private Sub Setting_Click()
  ConfigScrn.Show
End Sub



Public Sub FirstDo()
  Set gsDBName = OpenDatabase("d:\hxg\comm.mdb")
  Set gsRecordSend = gsDBName.OpenRecordset("send_table", dbOpenDynaset)
  Set gsRecordComm = gsDBName.OpenRecordset("comm_table", dbOpenDynaset)
  Dim Shared Nbaud$, Nparity$
  Dim Shared Ndata$, NStop$
  Dim Shared Nsettings
  Select Case gsRecordComm.Fields("port").Value
    Case "com1":
      MSComm1.Port = com1
    Case "com1":
      MSComm1.Port = com2
  End Select
  Nbaud$ = gsRecordComm.Fields("bound").Value
  Nparity$ = gsRecordComm.Fields("parity").Value
  Ndata$ = gsRecordComm.Fields("data").Value
  NStop$ = gsRecordComm.Fields("stop").Value
  Nsettings = Nbaud$ + "," + Nparity$ + "," + Ndata$ + "," + NStop$
  MSComm1.Settings = Nsettings
  MSComm1.Handshaking = gsRecordComm.Fields("control").Value
  
End Sub

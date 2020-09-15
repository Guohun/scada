VERSION 4.00
Begin VB.Form transmitForm 
   Caption         =   "Form2"
   ClientHeight    =   1290
   ClientLeft      =   1035
   ClientTop       =   2475
   ClientWidth     =   6690
   Height          =   1695
   Left            =   975
   LinkTopic       =   "Form2"
   ScaleHeight     =   1290
   ScaleWidth      =   6690
   Top             =   2130
   Width           =   6810
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5280
      Top             =   120
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
Attribute VB_Name = "transmitForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

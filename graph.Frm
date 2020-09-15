VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   1380
   ClientTop       =   1710
   ClientWidth     =   6690
   Height          =   6150
   Left            =   1320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   6690
   Top             =   1080
   Width           =   6810
   Begin VB.OLE OLE1 
      Class           =   "Excel.Chart.5"
      Height          =   1815
      Left            =   810
      OleObjectBlob   =   "graph.frx":0000
      TabIndex        =   0
      Top             =   1410
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub OLE1_Updated(Code As Integer)

End Sub


Private Sub OLE2_Updated(Code As Integer)

End Sub



VERSION 4.00
Begin VB.Form NewRec 
   Caption         =   "NewRec Form:"
   ClientHeight    =   2145
   ClientLeft      =   1500
   ClientTop       =   1665
   ClientWidth     =   6720
   Height          =   2550
   Left            =   1440
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6720
   Top             =   1320
   Width           =   6840
   Begin VB.CommandButton Delete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Edit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton AddNew 
      Caption         =   "&Add New"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Data NewDB 
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid NDBGrid1 
      Bindings        =   "NewRec.frx":0000
      Height          =   1815
      Left            =   0
      OleObjectBlob   =   "NewRec.frx":000E
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "NewRec"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub DBGrid1_Click()

End Sub

Private Sub AddNew_Click()
  NDBGrid1.AllowAddNew = True
  NDBGrid1.SetFocus
End Sub

Private Sub Delete_Click()
  NDBGrid1.AllowDelete = True
  NDBGrid1.SetFocus
End Sub

Private Sub Edit_Click()
  NDBGrid1.AllowUpdate = True
  NDBGrid1.SetFocus
End Sub


Private Sub Form_Load()
   newdb.DatabaseName = "c:\ljxt\ljxt.mdb"
   newdb.RecordSource = "lian_jiao_table"
End Sub


Private Sub Form_Resize()
  NDBGrid1.Width = NewRec.Width
  NDBGrid1.Height = NewRec.Width - 375
End Sub



VERSION 4.00
Begin VB.Form frmMdi1 
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   1185
   ClientTop       =   3075
   ClientWidth     =   6480
   Height          =   2415
   Left            =   1125
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2010
   ScaleWidth      =   6480
   Top             =   2730
   Width           =   6600
   Begin VB.Data Data2 
      Align           =   1  'Align Top
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   6480
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   2055
      Left            =   0
      OleObjectBlob   =   "xinliao.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmMdi1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

VERSION 4.00
Begin VB.Form frmMdi2 
   Caption         =   "Form2"
   ClientHeight    =   2040
   ClientLeft      =   2415
   ClientTop       =   3780
   ClientWidth     =   6480
   Height          =   2445
   Left            =   2355
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   6480
   Top             =   3435
   Width           =   6600
   Begin VB.Data Data1 
      Align           =   1  'Align Top
      Caption         =   "Data1"
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
      OleObjectBlob   =   "jiliao.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmMdi2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

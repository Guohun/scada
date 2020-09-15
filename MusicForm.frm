VERSION 4.00
Begin VB.Form MusicForm 
   Caption         =   "Form2"
   ClientHeight    =   855
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   Height          =   1260
   Left            =   1080
   LinkTopic       =   "Form2"
   ScaleHeight     =   855
   ScaleWidth      =   6690
   Top             =   1170
   Width           =   6810
End
Attribute VB_Name = "MusicForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
End Sub


Private Sub Form_Terminate()

End Sub


Private Sub Form_Unload(Cancel As Integer)
  MMControl1.Command = "close"
End Sub



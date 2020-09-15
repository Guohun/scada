VERSION 4.00
Begin VB.MDIForm TMDIForm 
   BackColor       =   &H8000000C&
   Caption         =   "Tan_Hei_Table_Input:"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   7005
   Height          =   6735
   Left            =   -15
   LinkTopic       =   "MDIForm1"
   Top             =   225
   Width           =   7125
End
Attribute VB_Name = "TMDIForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
  'BaseRec.Show
  'NewRec.Show
  If china_english = True Then Me.Caption = "Ì¿ºÚÅä·½Â¼Èë:"
  BaseRec.Show
  Tan_hei_table_Input.Show
  'MDIForm1.Arrange 1
End Sub


Private Sub MDIForm_Resize()
Tan_hei_table_Input.Top = 0
Tan_hei_table_Input.Width = Tan_hei_table_Input.Width
BaseRec.Width = Tan_hei_table_Input.Width
BaseRec.Top = Tan_hei_table_Input.Height + Tan_hei_table_Input.Top
End Sub



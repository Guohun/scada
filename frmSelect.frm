VERSION 4.00
Begin VB.Form frmSelect 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pei_fang_table"
   ClientHeight    =   516
   ClientLeft      =   2592
   ClientTop       =   1560
   ClientWidth     =   2568
   Height          =   900
   Left            =   2544
   LinkTopic       =   "Form1"
   ScaleHeight     =   516
   ScaleWidth      =   2568
   Top             =   1224
   Width           =   2664
   Begin VB.ComboBox CMBPf 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "txtPei_fang"
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim gsDPname As Database
Dim gsDPrecordset As Recordset
Dim scount As Integer

Private Sub CMBPf_Click()
Dim i As Integer
For i = 0 To scount - 1
  If i = CMBPf.ListIndex Then Exit For
  gsDPrecordset.MoveNext
  
Next
produce!Grid1.Col = 1
produce!Grid1.Text = gsDPrecordset.Fields("number").Value
produce!Grid1.Col = 2
produce!Grid1.Text = gsDPrecordset.Fields("name").Value
gsDPrecordset.MoveFirst
'gsDPrecordset.Close
'gsDPname.Close
End Sub


Private Sub CMBPf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Unload Me
End Sub


Private Sub Form_Load()
If china_english = True Then Me.Caption = "Åä·½Ä¿Â¼±í"
Set gsDPname = OpenDatabase("c:\ljxt\ljxt.mdb")
Set gsDPrecordset = gsDPname.OpenRecordset("pei_fan_table", dbOpenDynaset)
gsDPrecordset.MoveLast
scount = gsDPrecordset.RecordCount
gsDPrecordset.MoveFirst
For i = 0 To scount - 1
  CMBPf.AddItem gsDPrecordset.Fields("number").Value + Space(15 - Len(gsDPrecordset.Fields("number").Value)) + gsDPrecordset.Fields("name").Value
  gsDPrecordset.MoveNext
Next
gsDPrecordset.MoveFirst
CMBPf.Text = CMBPf.List(CMBPf.ListIndex)
End Sub



VERSION 4.00
Begin VB.Form BaseRec 
   Caption         =   "BaseRec Form:"
   ClientHeight    =   1650
   ClientLeft      =   1470
   ClientTop       =   5490
   ClientWidth     =   6690
   Height          =   2055
   Left            =   1410
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6690
   Top             =   5145
   Width           =   6810
   Begin VB.CommandButton SelectRec 
      Caption         =   "&Select A Record"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   6735
   End
   Begin VB.Data BaseDB 
      Caption         =   "BaseDB"
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid BDBGrid1 
      Bindings        =   "BaseRec.frx":0000
      Height          =   1215
      Left            =   0
      OleObjectBlob   =   "BaseRec.frx":000F
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "BaseRec"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub Form_Load()
  BaseDB.DatabaseName = "c:\ljxt\ljxt.mdb"
  BaseDB.RecordSource = "cai_liao_table"
  Me.Top = 4270
  Me.Height = 2000
  Me.Left = 0
End Sub


Private Sub Form_Resize()
  'BDBGrid1.Width = BaseRec.Width
  'BDBGrid1.Height = BaseRec.Height - 500 '375
  'SelectRec.Top = BDBGrid1.Height - 370
  'SelectRec.Width = BaseRec.Width
End Sub





Private Sub SelectRec_Click()
  Dim FldName As Field
  Dim i, j As Integer
  If (MsgBox("Select Current Record to DTan_hei_table?", 4)) <> 6 Then Exit Sub
  For i = 0 To BaseDB.Recordset.Fields.count - 1
  For Each FldName In BaseDB.Recordset.Fields
  For j = 0 To Tan_hei_table_Input.DTan_hei_table.Recordset.Fields.count - 1
   If (FldName.Name = Tan_hei_table_Input.DTan_hei_table.Recordset.Fields(i).Name) Then
    'NewRec.NDBGrid1.Columns(1).Value = BaseDB.Recordset.Fields(i).Value
     Tan_hei_table_Input.DTan_hei_table.Recordset.Fields(i).Value = FldName.Value
   End If
  Next
  Next
  Next
  If (MsgBox("Insert The New Recordset?", 4)) <> 6 Then NewRec.newdb.Recordset.CancelUpdate
'  NewRec.newdb.Recordset.Update
'  NewRec.NDBGrid1.Refresh
End Sub



VERSION 4.00
Begin VB.Form frmTables 
   Caption         =   "Tables"
   ClientHeight    =   3105
   ClientLeft      =   2040
   ClientTop       =   2325
   ClientWidth     =   3735
   Height          =   3795
   HelpContextID   =   2016146
   Icon            =   "TABLES.frx":0000
   Left            =   1980
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleMode       =   0  'User
   ScaleWidth      =   3735
   Top             =   1695
   Width           =   3855
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstTables 
      BackColor       =   &H00FFFFFF&
      Height          =   1620
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Private Sub cmdOpen_Click()
'Dim a As String
'a = lstQueryDefs.Text 'List(lstQueryDefs.ListIndex)
'Text1.Text = a
'Set mrecFormRecordset = gsDBName.OpenRecordset(StripConnect(a), dbOpenDynaset)
Hide

frmDynaSnap.Show

End Sub

Private Sub Form_Load()
Set gsDBName = OpenDatabase("c:\ljxt\ljxt.mdb")
Me.Caption = "DBName:" & "ljxt.mdb"
gsDBName.TableDefs.Refresh
RefreshTables frmTables.lstTables, True
china_english = False
gsRecordSet = lstTables.Text
password = "pass"
user_name = "exe"
End Sub

Private Sub lstQueryDefs_Click()
gsRecordSet = lstQueryDefs.Text
End Sub

Private Sub lstQueryDefs_DblClick()
gsRecordSet = lstQueryDefs.Text
Call cmdOpen_Click
End Sub


Private Sub lstTables_Click()
gsRecordSet = lstTables.Text
End Sub


Private Sub lstTables_DblClick()
gsRecordSet = lstTables.Text
Call cmdOpen_Click
End Sub


Private Sub mnuFileItem_Click(Index As Integer)
Select Case Index
  Case 0:
    Unload Me
    End
End Select
    
  
End Sub



VERSION 4.00
Begin VB.Form Statistic 
   Caption         =   "Statistic"
   ClientHeight    =   1335
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   6105
   Height          =   1740
   Left            =   1830
   LinkTopic       =   "form1"
   ScaleHeight     =   1335
   ScaleWidth      =   6105
   Top             =   1560
   Width           =   6225
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\hxg\produce\statistic.mdb"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "dayStatistic"
      Top             =   2040
      Width           =   3855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Statistest2.frx":0000
      Height          =   1335
      Left            =   0
      OleObjectBlob   =   "Statistest2.frx":000E
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Statistic"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim InDate As String
Dim InBan As String
Dim china_english As Integer
Dim recRecordset1 As Recordset, recRecordset2 As Recordset
Public Sub InitData()
'Data1.Recordset.MoveLast
gcount = Data1.Recordset.RecordCount
If gcount = 0 Then GoTo DRefresh
Data1.Recordset.MoveFirst
For i = 0 To gcount - 1
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
Next
Data1.Recordset.MoveFirst
DRefresh:
 Dim gsDBName As Database
 Dim gsRecord As Recordset
 Dim gsTRecord As Recordset
 Set gsDBName = OpenDatabase("c:\ljxt\make.mdb")
 Set gsRecord = gsDBName.OpenRecordset("make_table", dbOpenDynaset)
' gsRecord.Filter = "[ban]='" + InBan + "'"
' Set gsRecord = gsRecord.OpenRecordset()
' gsRecord.Filter = "[date]='" + InDate + "'"
' Set gsRecord = gsRecord.OpenRecordset()
 Set gsTRecord = gsDBName.OpenRecordset("make_pei_fan", dbOpenDynaset)
 gsTRecord.Filter = "[only_key]=" + Str(gsRecord.Fields("only_key").Value) '+ "'"
Set gsTRecord = gsTRecord.OpenRecordset()
gsTRecord.MoveLast
tcount = gsTRecord.RecordCount
gsTRecord.MoveFirst
For i = 0 To tcount - 1
Data1.Recordset.AddNew
Data1.Recordset.Fields("ban").Value = gsTRecord.Fields("machine").Value
Data1.Recordset.Fields("name").Value = gsRecord.Fields("name").Value
Data1.Recordset.Fields("che").Value = gsRecord.Fields("che").Value
Data1.Recordset.Fields("weightneed").Value = gsTRecord.Fields("w1").Value
Data1.Recordset.Fields("trueweight").Value = gsTRecord.Fields("w2").Value
Data1.Recordset.Fields("surplus").Value = gsTRecord.Fields("w1").Value - gsTRecord.Fields("w2").Value
Data1.Recordset.Update
gsTRecord.MoveNext
Next
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
Select Case ColIndex
  Case 0
    Set Data1.Recordset = recRecordset2
  Case 4
    Set Data1.Recordset = recRecordset1
    DBGrid1.Refresh
End Select
End Sub


Private Sub Form_Load()
china_english = True
'InitData


End Sub


Private Sub Form_Resize()
If china_english = True Then
    DBGrid1.Columns(0).Caption = "机组号"
    DBGrid1.Columns(1).Caption = "配方名"
    DBGrid1.Columns(2).Caption = "车数"
    DBGrid1.Columns(3).Caption = "材料所需量"
    DBGrid1.Columns(4).Caption = "实际量"
    DBGrid1.Columns(5).Caption = "余量"
End If
InitData
SortData
'DBGrid1.Width = Me.Width
'DBGrid1.RowHeight = DBGrid1.RowHeight * Me.Height / DBGrid1.Height
''DBGrid1.colwidth = DBGrid1.colwidth * Me.Width / DBGrid1.Width
'DBGrid1.Height = Me.Height

End Sub



Public Sub SortData()
  'On Error GoTo SortErr
  'Data1.Recordset.MoveFirst
  Data1.Recordset.Sort = Data1.Recordset.Fields(4).Name 'sortname 'SortStr
  Set recRecordset1 = Data1.Recordset.OpenRecordset() 'recRecordset1.Type)
  Data1.Recordset.Sort = Data1.Recordset.Fields(0).Name 'sortname 'SortStr
  Set recRecordset2 = Data1.Recordset.OpenRecordset() 'recRecordset1.Type)
  Data1.Recordset.Sort = Data1.Recordset.Fields(0).Name 'sortname 'SortStr
  Set Data1.Recordset = Data1.Recordset.OpenRecordset() 'recRecordset1.Type)
  Exit Sub

SortErr:
  'Screen.MousePointer = 0
  MsgBox "Error:" & Err & " " & Error$
  Exit Sub

End Sub

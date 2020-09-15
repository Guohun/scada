VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1005
   ClientLeft      =   2490
   ClientTop       =   2400
   ClientWidth     =   4725
   Height          =   1410
   Left            =   2430
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   4725
   Top             =   2055
   Width           =   4845
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\hxg\produce\statistic.mdb"
      Exclusive       =   0   'False
      Height          =   495
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "dayStatistic"
      Top             =   3360
      Width           =   4455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "test.frx":0000
      Height          =   855
      Left            =   0
      OleObjectBlob   =   "test.frx":000E
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim InDate As String
Dim InBan As String
Dim china_english As Integer






Private Sub DBGrid1_Click()
'MsgBox DBGrid1.Columns(0).Caption
'DBGrid1.Columns(0).Caption = "中国"
If china_english = True Then
    DBGrid1.Columns(0).Caption = "机组号"
    DBGrid1.Columns(1).Caption = "配方名"
    DBGrid1.Columns(2).Caption = "车数"
    DBGrid1.Columns(3).Caption = "材料所需量"
    DBGrid1.Columns(4).Caption = "实际量"
    DBGrid1.Columns(5).Caption = "余量"
End If
End Sub

Public Sub InitData()
Data1.MoveLast
gcount = Data1.Recordset.RecordCount
If gcount = 0 Then GoTo DRefresh
Data1.MoveFirst
For i = 0 To gcount - 1
  Data1.Delete
  Data1.MoveNext
Next
Data1.MoveFirst
DRefresh:
 Dim gsDBName As Database
 Dim gsRecord As Recordset
 Dim gsTRecord As Recordset
 Set gsDBName = OpenDatabase("c:\ljxt\make.mdb")
 Set gsRecord = gsDBName.OpenRecordset("make_table", dbOpenDynaset)
 gsRecord.Filter = "[ban]='" + InBan + "'"
 Set gsRecord = gsDBName.OpenRecordset()
 gsRecord.Filter = "[date]='" + InDate + "'"
 Set gsRecord = gsDBName.OpenRecordset()
 Set gsTRecord = gsDBName.OpenRecordset("make_pei_fan", dbOpenDynaset)
 Set gsTRecord.Filter = "[only_key]='" + gsrecordset.Fields("only_key").Value + "'"
gsTRecord = gsDBName.OpenRecordset()
Data1.Recordset.AddNew
Data1.Recordset.Fields("ban").Value = gsRecord.Fields("ban").Value
Data1.Recordset.Fields("name").Value = gsRecord.Fields("name").Value
Data1.Recordset.Fields("che").Value = gsRecord.Fields("che").Value
Data1.Recordset.Fields("weightneed").Value = gsTRecord.Fields("w1").Value
Data1.Recordset.Fields("trueweight").Value = gsTRecord.Fields("w2").Value
Data1.Recordset.Fields("surplus").Value = gsTRecord.Fields("w1").Value - gsTRecord.Fields("w2").Value
Data1.Update
DBGrid1.Refresh
End Sub


Private Sub Form_Load()
china_english = True
End Sub


Private Sub Form_Resize()
DBGrid1.Top = DBGrid1.Top - 100
Call DBGrid1_Click
End Sub



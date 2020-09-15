VERSION 4.00
Begin VB.Form StatisticOnDay 
   Caption         =   "Statistic2"
   ClientHeight    =   2445
   ClientLeft      =   1845
   ClientTop       =   3090
   ClientWidth     =   5820
   Height          =   2850
   Left            =   1785
   LinkTopic       =   "Statis2"
   ScaleHeight     =   2445
   ScaleWidth      =   5820
   Top             =   2745
   Width           =   5940
   Begin VB.TextBox txtBan 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\make.Mdb"
      Exclusive       =   0   'False
      Height          =   495
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "day_statistic"
      Top             =   3360
      Width           =   4455
   End
   Begin Threed.SSPanel Title 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
   End
   Begin Threed.SSPanel Datelist 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   " "
      BackColor       =   12632256
   End
   Begin Threed.SSPanel Reporter 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   " "
      BackColor       =   12632256
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "SSCommand1"
      Enabled         =   0   'False
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "StatisticOnDay.frx":0000
      Height          =   1575
      Left            =   0
      OleObjectBlob   =   "StatisticOnDay.frx":000E
      TabIndex        =   0
      Top             =   840
      Width           =   5775
   End
End
Attribute VB_Name = "StatisticOnDay"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim InDate, InDate2 As String
Dim InBan As String
Dim china_english As Integer
Dim user_name, department As String

Public Sub InitData()
On Error GoTo ErrD
Data1.Recordset.MoveLast
NoRecord:
Dim gcount As Integer
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
 Dim gsRecordBak As Recordset
 Dim gsTRecord As Recordset
 Set gsDBName = OpenDatabase("c:\ljxt\make.mdb")
 Set gsRecord = gsDBName.OpenRecordset("make_table", dbOpenDynaset)
' gsRecord.Filter = "[ban]='" + InBan + "'"
' Set gsRecord = gsDBName.OpenRecordset()
Date = InDate
gsRecord.Filter = "[_date]> #" & Format(Date, "yy-m-d") & "#" '+ "'" + Date + "'" 'InDate,,)
Set gsRecord = gsRecord.OpenRecordset()
Date = InDate2
gsRecord.Filter = "[_date]< #" & Format(Date, "yy-m-d") & "#" '+ "'" + Date + "'" 'InDate,,)
gsRecord.MoveLast
gcount = gsRecord.RecordCount - 1
gsRecord.MoveFirst
BCount = GetRecordInfo(gsRecord, gcount)
For k = 0 To BCount
'过滤班----统计班表
  Set gsRecordBak = gsRecord
  gsRecordBak.Filter = "[ban]=" + "'" + Val(txtBan(i).Text) + "'"
  Set gsRecordBak = gsRecordBak.OpenRecordset()
  Data1.Recordset.AddNew
  Data1.Recordset.Fields("ban").Value = gsRecordBak.Fields("ban").Value
  Data1.Recordset.Fields("name").Value = gsRecordBak.Fields("name").Value
  gsRecordBak.MoveLast
  gbcount = gsRecordBak.RecordCount - 1
  gsRecordBak.MoveFirst
  For i = 0 To gbcount
    Set gsTRecord = gsDBName.OpenRecordset("make_pei_fan", dbOpenDynaset)
    gsTRecord.Filter = "[only_key]=" + Str(gsRecordBak.Fields("only_key").Value)
    Set gsTRecord = gsTRecord.OpenRecordset()
    If gsTRecord.EOF Then
      MsgBox "No Match Record!"
      Exit Sub
    End If
    gsTRecord.MoveLast
    gscount = gsTRecord.RecordCount - 1
    gsTRecord.MoveFirst
    chenumber = 0
    wastenumber = 0
    For j = 0 To gscount
        chenumber = chenumber + Val(gsTRecord.Fields("now_che").Value)
        'Data1.Recordset.Fields("weightneed").Value = gsTRecord.Fields("w1").Value
        wastenumber = Val(gsTRecord.Fields("w2").Value) + wastenumber
    Next
    Data1.Recordset.Fields("now_che").Value = Str(chenumber)
    Data1.Recordset.Fields("waste").Value = Str(wastenumber)
    Data1.Recordset.Update
    Data1.Recordset.MoveFirst
    gsRecordBak.MoveNext
Next
Next
'DBGrid1.Refresh
Exit Sub
ErrD:
  If Err = 3021 Then GoTo NoRecord
End Sub





Private Sub Datelist_Click()
InDate = InputBox$("please input the first date:")
InDate2 = InputBox$("Please Input The Second Date:")
If InDate <> "" And InDate2 <> "" Then InitData
End Sub

Private Sub Form_Load()
china_english = True
InDate = InputBox$("please input the first date:")
InDate2 = InputBox$("Please Input The Second Date:")
user_name = "桂林"
department = "电子工业学院"
End Sub


Private Sub Form_Resize()
DBGrid1.Top = DBGrid1.Top - 100
If china_english = True Then
    DBGrid1.Columns(0).Caption = "班号"
    DBGrid1.Columns(1).Caption = "配方名"
    DBGrid1.Columns(2).Caption = "当前车数"
'    DBGrid1.Columns(3).Caption = "材料所需量"
    DBGrid1.Columns(3).Caption = "消耗量"
'    DBGrid1.Columns(5).Caption = "余量"
Title.Caption = department + "单位日材料统计"
Datelist.Caption = Left(InDate, 2) + "年" + Right(Left(InDate, 5), 2) + "月" + Right(InDate, 2) + "日"
Reporter.Caption = user_name + "打印"
End If
Title.Caption = department + "单位日材料统计"
Datelist.Caption = Left(InDate, 2) + "年" + Right(Left(InDate, 5), 2) + "月" + Right(InDate, 2) + "日"
Reporter.Caption = user_name + "打印"
If InDate <> "" And InDate2 <> "" Then InitData
End Sub



Private Function GetRecordInfo(gRecord As Recordset, gscount As Integer)
txtBan(0).Text = gRecord.Fields("ban").Value
gRecord.MoveNext
j = 0
For i = 1 To gscount
  If gRecord.Fields("ban").Value <> txtBan(j) Then
   j = j + 1
   Load txtBan(j)
   txtBan(j).Text = gRecord.Fields("ban").Value
  End If
  gRecord.MoveNext
Next
gRecord.MoveFirst
GetRecordInfo = j
End Function

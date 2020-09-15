VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   2850
   ClientTop       =   2625
   ClientWidth     =   3870
   Height          =   1260
   Left            =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   3870
   Top             =   2280
   Width           =   3990
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   3855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   1508
      _StockProps     =   15
      Caption         =   "Jiao_liao"
      TabsPerRow      =   3
      Tab             =   2
      TabOrientation  =   0
      Tabs            =   3
      Style           =   0
      TabMaxWidth     =   0
      TabHeight       =   529
      TabCaption(0)   =   "Tan_Hei"
      Tab(0).ControlCount=   0
      Tab(0).ControlEnabled=   0   'False
      TabCaption(1)   =   "You_Liao"
      Tab(1).ControlCount=   0
      Tab(1).ControlEnabled=   0   'False
      TabCaption(2)   =   "Jiao_liao"
      Tab(2).ControlCount=   0
      Tab(2).ControlEnabled=   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim gsSelectR As Recordset
Dim DBName As Database

Private Sub Combo1_Click()
Dim inturn As Integer
inturn = 0
combo1.ListIndex = 0
gsSelectR.MoveFirst
Do
If inturn = combo1.ListIndex Then Exit Do
inturn = inturn + 1
Loop
If UseCombo1 = True Then
lian_jiao_table_Input!txtFields(1).Text = gsThisSet.Fields("number").Value
Else
lian_jiao_table_Input!txtFields(2).Text = gsThisSet.Fields("number").Value
End If
End Sub


Private Sub Form_Load()
Dim count As Integer
If china_english = True Then
  For i = 0 To 2
    Select Case i
     Case 0
       SSTab1.TabCaption(i) = "Ãø∫⁄"
     Case 1
       SSTab1.TabCaption(i) = "Ω∫¡œ"
     Case 2
       SSTab1.TabCaption(i) = "”Õ¡œ"
    End Select
  Next
End If
Set DBName = OpenDatabase("c:\ljxt\ljxt.mdb")
SSTab1.TabIndex = 0
Call SSTab1_Click(SSTab1.TabIndex)
'Set gsSelectR = DBName.OpenRecordset("tan_hei_table", dbOpenDynaset)
'gsSelectR.MoveLast
'count = gsSelectR.RecordCount
'gsSelectR.MoveFirst
'For i = 0 To count - 1
'  Combo1.AddItem "   " + Str(gsSelectR.Fields("number").Value) + Space(20) + gsSelectR.Fields("name").Value '+ Space(15 - Len(gsSelectR.Fields("name").Value)) ' + gsSelectR.Fields("dou").Value
'  gsSelectR.MoveNext
'Next
'gsSelectR.MoveFirst
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Dim count As Integer
Select Case PreviousTab
  Case 0:
combo1.Clear
Set gsSelectR = DBName.OpenRecordset("tan_hei_table", dbOpenDynaset)
gsSelectR.MoveLast
count = gsSelectR.RecordCount
gsSelectR.MoveFirst
For i = 0 To count - 1
  combo1.AddItem "  0 " + Str(gsSelectR.Fields("number").Value) + Space(20) + gsSelectR.Fields("name").Value '+ Space(15 - Len(gsSelectR.Fields("name").Value)) ' + gsSelectR.Fields("dou").Value
  gsSelectR.MoveNext
Next
gsSelectR.MoveFirst
  Case 1:
combo1.Clear
Set gsSelectR = DBName.OpenRecordset("jiao_liao_table", dbOpenDynaset)
gsSelectR.MoveLast
count = gsSelectR.RecordCount
gsSelectR.MoveFirst
For i = 0 To count - 1
  combo1.AddItem " 2  " '+ gsSelectR.Fields("number").Value + Space(20) + gsSelectR.Fields("name").Value + Space(15 - Len(gsSelectR.Fields("name").Value)) + gsSelectR.Fields("dou").Value
  gsSelectR.MoveNext
Next
gsSelectR.MoveFirst
  Case 2:
combo1.Clear
Set gsSelectR = DBName.OpenRecordset("you_liao_table", dbOpenDynaset)
gsSelectR.MoveLast
count = gsSelectR.RecordCount
gsSelectR.MoveFirst
For i = 0 To count - 1
  combo1.AddItem " 1  " '+ gsSelectR.Fields("number").Value + Space(20) + gsSelectR.Fields("name").Value + Space(15 - Len(gsSelectR.Fields("name").Value)) + gsSelectR.Fields("dou").Value
  gsSelectR.MoveNext
Next
gsSelectR.MoveFirst
End Select
End Sub


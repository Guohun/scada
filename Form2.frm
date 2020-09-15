VERSION 4.00
Begin VB.Form produce 
   Caption         =   "Produce"
   ClientHeight    =   2295
   ClientLeft      =   1815
   ClientTop       =   2175
   ClientWidth     =   5475
   Height          =   2985
   Left            =   1755
   LinkTopic       =   "Form2"
   ScaleHeight     =   2295
   ScaleWidth      =   5475
   Top             =   1545
   Width           =   5595
   Begin MSGrid.Grid Grid1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   3836
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   14737632
      Rows            =   100
      Cols            =   7
      FixedCols       =   0
   End
   Begin VB.Menu mnuEnd 
      Caption         =   "&End"
      Begin VB.Menu mnuEndItems 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEndItems 
         Caption         =   "&Exit"
         Index           =   1
      End
   End
End
Attribute VB_Name = "produce"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Form_Load()
 china_english = True
 run_flag = 1
 FirstDo
End Sub


Public Sub FirstDo()
  Grid1.ColWidth(0) = 800
  Grid1.ColWidth(1) = 800
  Grid1.ColWidth(2) = 1500
  Grid1.ColWidth(3) = 400
  Grid1.ColWidth(4) = 400
  Grid1.ColWidth(5) = 400
  Grid1.ColWidth(6) = 800
  For i = 0 To 99
    Grid1.RowHeight(i) = 320
  Next
  Grid1.Row = 0
  If china_english = True Then
    Grid1.Col = 0
    Grid1.Text = "目录序号"
    Grid1.Col = 1
    Grid1.Text = "配方编号"
    Grid1.Col = 2
    Grid1.Text = "配方名称"
    Grid1.Col = 3
    Grid1.Text = "班号"
    Grid1.Col = 4
    Grid1.Text = "车数"
    Grid1.Col = 5
    Grid1.Text = "标志"
    Grid1.Col = 6
    Grid1.Text = "当前车数"
 Else
    Grid1.Col = 0
    Grid1.Text = "index"
    Grid1.Col = 1
    Grid1.Text = "number"
    Grid1.Col = 2
    Grid1.Text = "name"
    Grid1.Col = 3
    Grid1.Text = "ban"
    Grid1.Col = 4
    Grid1.Text = "che"
    Grid1.Col = 5
    Grid1.Text = "flag"
    Grid1.Col = 6
    Grid1.Text = "now_che"
End If
Grid1.Row = 1
Grid1.Col = 0
Grid1.Text = 1
Grid1.Col = 5
Grid1.Text = 0
Grid1.Col = 6
Grid1.Text = 0
End Sub

Private Sub Grid1_Click()
If run_flag = 1 Then Exit Sub
bcol = Grid1.Col
brow = Grid1.Row
Grid1.Col = 0
If Grid1.Text = "" And IsLastFull = True And Grid1.Row > 1 Then
   Grid1.Row = brow - 1
   tnumber = Val(Grid1.Text)
   Grid1.Row = brow
   Grid1.Text = Str(tnumber + 1)
   Grid1.Col = 5
   Grid1.Text = 0
   Grid1.Col = 6
   Grid1.Text = 0
End If

Grid1.Col = bcol
End Sub

Private Sub Grid1_DblClick()
Col = Grid1.Col
Grid1.Col = 5
If Val(Grid1.Text) <> 0 Then
Grid1.Col = Col
Exit Sub
End If
Grid1.Col = 0
If Grid1.Text = "" Then
Grid1.Col = Col
Exit Sub
End If
Grid1.Col = Col
If Grid1.Col = 1 Or Grid1.Col = 2 Then frmSelect.Show vbModal
End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
Col = Grid1.Col
Grid1.Col = 5
If run_flag = 1 And Val(Grid1.Text) <> 0 Then
Grid1.Col = Col
Exit Sub
End If
Grid1.Col = Col
If KeyAscii = 13 And IsLastFull = True And IsCurrFull = False Then Call Grid1_Click
If (Grid1.Col = 1 Or Grid1.Col = 2) Then
If KeyAscii = 13 Then
Call Grid1_DblClick
Exit Sub
End If
End If
Grid1.Col = 0
If Grid1.Text = "" Then
Grid1.Col = Col
Exit Sub
End If
Grid1.Col = Col
If (Grid1.Col = 3 Or Grid1.Col = 4) Then
  If KeyAscii = 8 And Grid1.Text <> "" Then
    Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
  ElseIf KeyAscii = 13 Then
    SendKeys "{right}"
  ElseIf KeyAscii >= 49 And KeyAscii <= 59 Then
  If (Len(Grid1.Text) < 3) Then
  oldchar = Grid1.Text
  Grid1.Text = oldchar + Chr(KeyAscii)
  Else: Exit Sub
  End If
  End If
End If

End Sub


Public Function IsLastFull() As Integer
Dim bcol, brow As Integer
bcol = Grid1.Col
brow = Grid1.Row
Grid1.Row = brow - 1
For i = 0 To 6
  Grid1.Col = i
  If Grid1.Text = "" Then
  IsLastFull = False
  Grid1.Col = bcol
  Grid1.Row = brow
  Exit Function
  End If
Next
Grid1.Col = bcol
Grid1.Row = brow
IsLastFull = True
  End Function

Public Function IsCurrFull() As Integer
Dim bcol As Integer
bcol = Grid1.Col
For i = 0 To 6
  Grid1.Col = i
  If Grid1.Text = "" Then
  IsCurrFull = False
  Grid1.Col = bcol
  Exit Function
  End If
Next
Grid1.Col = bcol
IsCurrFull = True
End Function

Private Sub mnuEditItems_Click(Index As Integer)

End Sub

Private Sub mnuEndItems_Click(Index As Integer)
  If Index = 1 Then 'Call Form_Unload(True)
    Unload Me
  End If
End Sub

Private Sub mnuItems_Click(Index As Integer)
If Index = 1 Then Call Form_unload()
End Sub



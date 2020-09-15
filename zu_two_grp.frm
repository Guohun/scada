VERSION 4.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4155
   ClientLeft      =   810
   ClientTop       =   1605
   ClientWidth     =   6690
   Height          =   4560
   Left            =   750
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   6690
   Top             =   1260
   Width           =   6810
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   13.5
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.ComboBox Cb 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6120
      Top             =   3360
   End
   Begin GraphLib.Graph Gra 
      Height          =   1815
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   3201
      _StockProps     =   96
      BorderStyle     =   1
      GraphType       =   4
      RandomData      =   1
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
   Begin GraphLib.Graph Gra 
      Height          =   1815
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   3201
      _StockProps     =   96
      BorderStyle     =   1
      GraphType       =   1
      RandomData      =   1
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Base 1
Public time_num As Integer
       Dim a(2, 100) As Integer
       Dim b(2) As String
       Dim c(100) As String


Private Sub Command1_Click()
      
      
   'gra(1).Visible = IIf(gra.Visible = True, False, True)
       
  '     End
End Sub

Private Sub Form_Load()
       Cb.AddItem "gra.grahtype= 4"
          Cb.AddItem "gra.grahtype= 2"
          Cb.AddItem "gra.grahtype= 6"
        Cb.AddItem "gra.grahtype= 8"
        Cb.AddItem "gra.grahtype=1"
        
        Cb.AddItem "gra.grahtype= 3"
        
       Cb.Text = "gra.grahtype= 3"
        
        
        
        
        
        gra(0).GridStyle = 3
       time_num = 0
       a(1, 1) = 45: a(1, 2) = 2: a(1, 3) = 68: a(1, 4) = 90: a(1, 5) = 23: a(1, 6) = 8
       
       a(2, 1) = 6: a(2, 2) = 70: a(2, 3) = 39: a(2, 4) = 67: a(2, 5) = 33: a(2, 6) = 90
       b(1) = "ÎÂ¶È": b(2) = "¹¦ÂÊ"
       c(1) = "1": c(2) = "2": c(3) = "3": c(4) = "4"
       gra(0).GraphType = Val(Right(Trim(Cb.Text), 2))
       gra(0).NumSets = 2
       gra(0).NumPoints = 8

       'gra.ThisPoint =
       gra(0).GraphTitle = "SHOW POWER TEMPRO OF GRAPHE!"
       gra(0).AutoInc = 1
          For i = 1 To 2
             For j = 1 To 6
          gra(0).GraphData = a(i, j)
        Next j
        Next i
        For i = 1 To 6
          gra(0).LabelText = c(i)
        Next i
          For i = 1 To 2
          gra(0).LegendText = b(i)
          Next i
          gra(0).Labels = 1
          gra(0).DrawMode = 2
'=======================================
       gra(1).GridStyle = 3
       gra(1).GraphType = Val(Right(Trim(Cb.Text), 2))
       gra(1).NumSets = 2
       gra(1).NumPoints = 8

       'gra.ThisPoint =
       gra(1).GraphTitle = "SHOW POWER TEMPRO OF GRAPHE!"
       gra(1).AutoInc = 1
          For i = 1 To 2
             For j = 1 To 6
          gra(1).GraphData = a(i, j)
        Next j
        Next i
        For i = 1 To 6
          gra(1).LabelText = c(i)
        Next i
          For i = 1 To 2
          gra(1).LegendText = b(i)
          Next i
          gra(1).Labels = 1
          gra(1).DrawMode = 2
          
End Sub

Private Sub List1_Click()

End Sub

Private Sub Timer1_Timer()
       
  If flag = 0 Then
     '  Timer1.Enabled = flase
      ElseIf flag = -1 Then
      End
      Else
   ' Timer1.Enabled = True
      
End If

Static x, y As Integer
Static which_gra As Integer
which_gra = IIf(which_gra = 0, 1, 0)

time_num = time_num + 1
x = Abs(Int(100 * Rnd))
y = Abs(Int(100 * Rnd))

For i = 1 To 7
    a(1, i) = a(1, i + 1)
    a(2, i) = a(2, i + 1)
Next i
a(1, 8) = x
a(2, 8) = y
'    gra.NumPoints = gra.NumPoints - 1

       'gra.ThisPoint =
       gra(which_gra).GraphType = Val(Right(Trim(Cb.Text), 2))
       'gra(which_gra).NumSets = 2
             
       
       'gra(which_gra).AutoInc = 1
          For i = 1 To 2
             For j = 1 To gra(which_gra).NumPoints
          gra(which_gra).GraphData = a(i, j)
        Next j
        Next i
        For i = 0 To gra(which_gra).NumPoints - 1
          gra(which_gra).LabelText = Str(time_num + i)
        Next i
          'For i = 1 To 2
          'gra.LegendText = b(i)
          'Next i
          gra(which_gra).Labels = 1
          gra(which_gra).DrawMode = 2
          gra(which_gra).ZOrder 0
          'gra(which_gra).Visible = True
          'If which_gra = 0 Then
          '   gra(1).Visible = False
          'Else
          '    gra(0).Visible = False
          'End If
End Sub

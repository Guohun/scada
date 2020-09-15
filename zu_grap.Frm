VERSION 4.00
Begin VB.Form bl 
   Caption         =   "Form2"
   ClientHeight    =   4155
   ClientLeft      =   945
   ClientTop       =   1740
   ClientWidth     =   6690
   Height          =   4560
   Left            =   885
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   6690
   Top             =   1395
   Width           =   6810
   Begin VB.ComboBox com2 
      BeginProperty Font 
         name            =   "宋体"
         charset         =   134
         weight          =   400
         size            =   14.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
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
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3600
      Width           =   2655
   End
   Begin VB.ComboBox Cb 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   0
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6240
      Top             =   720
   End
   Begin GraphLib.Graph Gra 
      Height          =   2715
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   4789
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
      Height          =   2715
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   4789
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
Attribute VB_Name = "bl"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Base 1
Public time_num As Integer
       Dim a(2, 100) As Integer
       Dim b(2) As String
       Dim c(100) As String
      Public flag As Integer
       


Private Sub com2_click()
        For i = 1 To 2
          gra(0).LegendText = b(i)
          Next i
           For i = 1 To 2
          gra(1).LegendText = b(i)
     
     Next i
        gra(0).Labels = 1
          gra(0).DrawMode = 2
        gra(1).Labels = 1
          gra(1).DrawMode = 2
       
       
     If com2.Text = "CHINA=中文" Then
         gra(0).GraphTitle = "展示温度——功率图!"
          gra(1).GraphTitle = "展示温度——功率图!"
           Command1.Caption = "退出"
    
        b(1) = "温度": b(2) = "功率"
        
        Cb.Clear
        Cb.AddItem "平面柱状图 = 3"
         Cb.AddItem "立体柱状图 = 4"
        Cb.AddItem "工程甘特图= 5"
        Cb.AddItem "业绩成长图 = 6"
         Cb.AddItem "业绩面积图 = 8"
        Cb.AddItem "业绩饼图 = 2"
        Cb.Text = "柱状图 = 4"
        
        
    Else
         gra(0).GraphTitle = "SHOW POWER TEMPRO OF GRAPHE!"
          gra(1).GraphTitle = "SHOW POWER TEMPRO OF GRAPHE!"
             Command1.Caption = "EXIT"
            
        b(1) = "TEMPRO": b(2) = "POWER"
       Cb.Clear
        Cb.AddItem "gra.grahtype= 4"
          Cb.AddItem "gra.grahtype= 2"
          Cb.AddItem "gra.grahtype= 6"
        Cb.AddItem "gra.grahtype= 8"
        Cb.AddItem "gra.grahtype=1"
        
        Cb.AddItem "gra.grahtype= 3"
        
       Cb.Text = "gra.grahtype= 3"
        
     End If
End Sub

Private Sub Command1_Click()
    End
      
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
        
        com2.AddItem "CHINA=中文"
        com2.AddItem "ENGLISH=英文"
       com2.Text = "ENGLISH=英文"
        
        gra(0).GridStyle = 3
       time_num = 0
       a(1, 1) = 45: a(1, 2) = 2: a(1, 3) = 68: a(1, 4) = 90: a(1, 5) = 23: a(1, 6) = 8
       
       a(2, 1) = 6: a(2, 2) = 70: a(2, 3) = 39: a(2, 4) = 67: a(2, 5) = 33: a(2, 6) = 90
      b(1) = "TEMPRO": b(2) = "POWER"
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
  ' If com2.Text = "CHINA=中文" Then
   '      gra(0).GraphTitle = "展示温度——功率图!"
    '      gra(1).GraphTitle = "展示温度——功率图!"
    '
    
     '   b(1) = "温度": b(2) = "功率"
   ' Else
    '     gra(0).GraphTitle = "SHOW POWER TEMPRO OF GRAPHE!"
     '     gra(1).GraphTitle = "SHOW POWER TEMPRO OF GRAPHE!"
    '
    
      '  b(1) = "TEMPRO": b(2) = "POWER"
    '
    ' End If
             
             
             
             
             
             
  If flag = 0 Then
    ' Exit Sub
      ElseIf flag = -1 Then
      End
      
   ' Timer1.Enabled = True
    ' End If
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
      '  For i = 1 To 6
       '   gra(0).LabelText = c(i)
       ' Next i
          For i = 1 To 2
          gra(0).LegendText = b(i)
          Next i
           For i = 1 To 2
          gra(1).LegendText = b(i)
          Next i
End Sub

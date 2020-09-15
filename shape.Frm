VERSION 4.00
Begin VB.Form Chart 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Chart"
   ClientHeight    =   4500
   ClientLeft      =   1215
   ClientTop       =   1125
   ClientWidth     =   8160
   FillColor       =   &H000000FF&
   FillStyle       =   4  'Upward Diagonal
   Height          =   4905
   Left            =   1155
   LinkTopic       =   "Form1"
   MouseIcon       =   "shape.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   8160
   Top             =   780
   Width           =   8280
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4440
      MouseIcon       =   "shape.frx":0442
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin Threed.SSPanel Ban 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   450
      _StockProps     =   15
      Caption         =   "SSPanel1"
      ForeColor       =   4194304
      BackColor       =   32896
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      FloodColor      =   16777152
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "100"
      BackColor       =   32896
      Font3D          =   1
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "200"
      BackColor       =   32896
      Font3D          =   1
   End
   Begin Threed.SSPanel SSPanel8 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "300"
      BackColor       =   32896
      Font3D          =   1
   End
   Begin Threed.SSPanel SSPanel9 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "400"
      BackColor       =   32896
      Font3D          =   1
   End
   Begin Threed.SSPanel SSPanel10 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "500"
      BackColor       =   32896
      Font3D          =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "SSCommand1"
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      FillStyle       =   5  'Downward Diagonal
      Height          =   1575
      Index           =   0
      Left            =   960
      Top             =   2160
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.Line Line16 
      X1              =   5760
      X2              =   5760
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line14 
      X1              =   4560
      X2              =   4560
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line10 
      X1              =   960
      X2              =   1080
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line9 
      X1              =   960
      X2              =   1080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line7 
      X1              =   1560
      X2              =   1560
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line5 
      X1              =   960
      X2              =   1080
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line4 
      X1              =   960
      X2              =   1080
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   960
      X2              =   1080
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   960
      Y1              =   3720
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   960
      X2              =   7200
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line12 
      X1              =   3360
      X2              =   3360
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line11 
      X1              =   2760
      X2              =   2760
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line8 
      X1              =   2160
      X2              =   2160
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line13 
      X1              =   3960
      X2              =   3960
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line15 
      X1              =   5160
      X2              =   5160
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line17 
      X1              =   6360
      X2              =   6360
      Y1              =   3720
      Y2              =   3600
   End
End
Attribute VB_Name = "Chart"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ErrC
'ForeColor = QBColor(5)
'DrawMode = 8
'For i = 1 To 10
'Load Shape(i)
'Shape(i).Height = 300 * i
'Shape(i).Width = 50 * i
'Shape(i).BackColor = QBColor(i + 3)
'Shape(i).Top = Line1.Y1 - Shape(i).Height
'Shape(i).Left = Line2.X1 + Shape(i).Width * (i)
'Line (Line2.X1, Shape(i).Top)-(Shape(i).Left, Shape(i).Top), , B
'Shape(i).Visible = True
'Next
'Exit Sub
Load Shape(2)
Shape(2).Visible = True
Shape(2).Top = Line1.Y1 - Val(test!Data1.Recordset.Fields("waste").Value) * 200
Shape(2).Height = Val(test!Data1.Recordset.Fields("waste").Value) * 200
Shape(2).Width = 300
Shape(2).Left = 400
ErrC:
  MsgBox "Object have been loaded!"
End Sub


Private Sub Command2_Click()
On Error GoTo ErrD
For i = 1 To 10
'Line (Line2.X1, Shape(i).Top)-(Shape(i).Left, Shape(i).Top), , BF
Unload Shape(i)
Next
Exit Sub
ErrD:
  MsgBox "Object have been unloaded!"
End Sub


Private Sub SSCommand1_Click()
'Statistic!Data1.Recordset.MoveLast
gcount = Statistic!Data1.Recordset.RecordCount
Statistic!Data1.Recordset.MoveFirst
If gcount = 0 Then
  MsgBox "No Data"
  Unload Me
End If
Shape(0).Height = Val(Statistic!Data1.Recordset.Fields("waste").Value) * 10
Shape(0).Top = Line1.Y1 - Shape(0).Height
Shape(0).Left = Line2.X1 + 200
Shape(0).Visible = True
Ban(0).Left = Shape(0).Left - 50
Ban(0).Caption = Statistic!Data1.Recordset.Fields("ban").Value + "°à"
Ban(0).Visible = True
For i = 1 To gcount - 1
  Statistic!Data1.Recordset.MoveNext
  Load Shape(i)
  Shape(i).Height = Val(Statistic!Data1.Recordset.Fields("waste").Value) * 10
  Shape(i).Top = Line1.Y1 - Shape(i).Height
  Shape(i).Left = Shape(i - 1).Left + 600
  Shape(i).BackColor = QBColor(i)
  Shape(i).Visible = True
  Load Ban(i)
  Ban(i).Left = Shape(i).Left - 50
  Ban(i).Caption = Statistic!Data1.Recordset.Fields("ban").Value + "°à"
  Ban(i).Visible = True
Next
End Sub



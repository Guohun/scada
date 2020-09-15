VERSION 5.00
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form F_Gd_Lc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "管道内径及颜色属性"
   ClientHeight    =   2880
   ClientLeft      =   1740
   ClientTop       =   2532
   ClientWidth     =   4752
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame2 
      Height          =   1692
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2172
      _Version        =   65536
      _ExtentX        =   3831
      _ExtentY        =   2985
      _StockProps     =   14
      Caption         =   "设置"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
      Begin VB.TextBox Text1 
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin Spin.SpinButton SpinButton1 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1200
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   661
         _StockProps     =   73
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   492
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   78
         Caption         =   "颜  色"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   4
      End
      Begin VB.Label Label1 
         Caption         =   "宽度："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   732
      End
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   852
      Left            =   2640
      TabIndex        =   10
      Top             =   1680
      Width           =   852
      _Version        =   65536
      _ExtentX        =   1503
      _ExtentY        =   1503
      _StockProps     =   78
      Caption         =   "确  定"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
      Picture         =   "f_gd_lc.frx":0000
   End
   Begin Threed.SSOption SSOption1 
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1572
      _Version        =   65536
      _ExtentX        =   2773
      _ExtentY        =   445
      _StockProps     =   78
      Caption         =   "供料管"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
   Begin Threed.SSOption SSOption1 
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1572
      _Version        =   65536
      _ExtentX        =   2773
      _ExtentY        =   445
      _StockProps     =   78
      Caption         =   "供油管"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
   Begin Threed.SSOption SSOption1 
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1572
      _Version        =   65536
      _ExtentX        =   2773
      _ExtentY        =   445
      _StockProps     =   78
      Caption         =   "供汽管"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   852
      Left            =   3600
      TabIndex        =   4
      Top             =   1680
      Width           =   852
      _Version        =   65536
      _ExtentX        =   1503
      _ExtentY        =   1503
      _StockProps     =   78
      Caption         =   "退  出"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
      Picture         =   "f_gd_lc.frx":0452
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   250
      X2              =   338
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Label Label2 
      Caption         =   "预览"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   2640
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "F_Gd_Lc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：演示时改变管道中流体的宽度及颜色

Dim tgd(0 To 2) As Gd_Type
Dim which_gd As Integer

Public Function demo_line(index As Integer)
Select Case index
    Case 0
      F_Gd_Lc.DrawWidth = Gqg.Width
      Line (Line1.X1, Line1.Y1)-(Line1.X2, Line1.Y2), Gqg.Color
      F_Gd_Lc.DrawWidth = tgd(0).Neijing_width
      Draw_Line F_Gd_Lc, Line1.X1, Line1.Y1, Line1.X2, Line1.Y2, Dls.Xd_len \ 2, Gqg.Color, tgd(0).Obj_color
      'gqg.obj_color = f_showlc!CommonDialog1.color
    Case 1
      F_Gd_Lc.DrawWidth = Gyg.Width
      Line (Line1.X1, Line1.Y1)-(Line1.X2, Line1.Y2), Gyg.Color
      F_Gd_Lc.DrawWidth = tgd(1).Neijing_width
      Draw_Line F_Gd_Lc, Line1.X1, Line1.Y1, Line1.X2, Line1.Y2, Dls.Xd_len \ 2, Gyg.Color, tgd(1).Obj_color
'      gyg.obj_color = f_showlc!CommonDialog1.color
    Case 2
      F_Gd_Lc.DrawWidth = Glg.Width
      Line (Line1.X1, Line1.Y1)-(Line1.X2, Line1.Y2), Glg.Color
      F_Gd_Lc.DrawWidth = tgd(2).Neijing_width
      Draw_Line F_Gd_Lc, Line1.X1, Line1.Y1, Line1.X2, Line1.Y2, Dls.Xd_len \ 2, Glg.Color, tgd(2).Obj_color
'      glg.obj_color = f_showlc!CommonDialog1.color
End Select
End Function

Private Sub Form_Activate()
    tgd(0).Neijing_width = Gqg.Neijing_width
    tgd(0).Obj_color = Gqg.Obj_color
    tgd(1).Neijing_width = Gyg.Neijing_width
    tgd(1).Obj_color = Gyg.Obj_color
    tgd(2).Neijing_width = Glg.Neijing_width
    tgd(2).Obj_color = Glg.Obj_color
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call Check_KeyCode(KeyCode, Shift)
End Sub

Private Sub Form_Load()
  Line1.BorderColor = Gqg.Color
  Line1.BorderWidth = Gqg.Width
End Sub

Private Sub SpinButton1_SpinDown()
If Val(Text1.Text) > 1 Then
    Text1.Text = str(Val(Text1.Text) - 1)
    tgd(which_gd).Neijing_width = Val(Text1.Text)
    demo_line which_gd
End If
End Sub

Private Sub SpinButton1_SpinUp()
Dim W As Integer
Select Case which_gd
    Case 0
        W = Gqg.Width
    Case 1
        W = Gyg.Width
    Case 2
        W = Glg.Width
End Select

If Val(Text1.Text) >= W Then
    Exit Sub
End If
Text1.Text = str(Val(Text1.Text) + 1)
tgd(which_gd).Neijing_width = Val(Text1.Text)
demo_line which_gd
End Sub

Private Sub SSCommand1_Click()
  F_ShowLc!CommonDialog1.ShowColor
  If SSOption1(0).Value = True Then
    tgd(0).Obj_color = F_ShowLc!CommonDialog1.Color
    demo_line 0
  End If
  If SSOption1(1).Value = True Then
    tgd(1).Obj_color = F_ShowLc!CommonDialog1.Color
    demo_line 1
  End If
  If SSOption1(2).Value = True Then
    tgd(2).Obj_color = F_ShowLc!CommonDialog1.Color
    demo_line 2
  End If
End Sub

Private Sub SSCommand2_Click()
  Hide
End Sub

Private Sub SSCommand3_Click()
Dim i As Integer
    Gqg.Neijing_width = tgd(0).Neijing_width
    Gqg.Obj_color = tgd(0).Obj_color
    Gyg.Neijing_width = tgd(1).Neijing_width
    Gyg.Obj_color = tgd(1).Obj_color
    Glg.Neijing_width = tgd(2).Neijing_width
    Glg.Obj_color = tgd(2).Obj_color
    
    For i = 0 To Zline_num - 1
        Call F_ShowLc.Hide_Zline(i, True)
    Next i
End Sub

Private Sub SSOption1_Click(index As Integer, Value As Integer)
 If Value = True Then
  Select Case index
    Case 0
        which_gd = 0
        demo_line 0
    Case 1
        which_gd = 1
        demo_line 1
    Case 2
        which_gd = 2
        demo_line 2
  End Select
  Text1.Text = str(tgd(which_gd).Neijing_width)
 End If
End Sub



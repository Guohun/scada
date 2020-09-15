VERSION 4.00
Begin VB.Form ComInput
   Caption         =   "ComInput"
   ClientHeight    =   4155
   ClientLeft      =   1440
   ClientTop       =   1425
   ClientWidth     =   6690
   Height          =   4560
   Left            =   1380
   LinkTopic       =   "ComInput"
   ScaleHeight     =   4155
   ScaleWidth      =   6690
   Top             =   1080
   Width           =   6810
   Begin VB.VScrollBar VScroll1 
      Height          =   3615
      LargeChange     =   10
      Left            =   6360
      Max             =   94
      TabIndex        =   14
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtNow_Che 
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "txtNow_che"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtFlag 
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "txtFlag"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtChe 
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   11
      Text            =   "txtChe"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtBan 
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   10
      Text            =   "txtBan"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Text            =   "txtName"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtNumber 
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   8
      Text            =   "txtNumber"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtIndex 
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "txtIndex"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Now_che"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Flag"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Che"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Ban"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Number"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Index"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "ComInput"
Attribute VB_Creatable = False

Attribute VB_Exposed = False
Private Sub DBGrid1_Click()

End Sub


Private Sub Form_Load()
For i = 1 To 100
 Load txtIndex(i)
 Load txtNumber(i)
 Load txtName(i)
 Load txtBan(i)
 Load txtChe(i)
 Load txtFlag(i)
 Load txtnow_che(i)
 txtIndex(i).Top = txtIndex(i - 1).Top + 390
 txtNumber(i).Top = txtNumber(i - 1).Top + 390
 txtName(i).Top = txtNumber(i - 1).Top + 390
 txtBan(i).Top = txtBan(i - 1).Top + 390
 txtChe(i).Top = txtChe(i - 1).Top + 390
 txtFlag(i).Top = txtFlag(i - 1).Top + 390
 txtnow_che(i).Top = txtnow_che(i - 1).Top + 390
 txtIndex(i).Visible = True
 txtNumber(i).Visible = True
 txtName(i).Visible = True
 txtBan(i).Visible = True
 txtChe(i).Visible = True
 txtFlag(i).Visible = True
 txtnow_che(i).Visible = True
 Next
End Sub


Private Sub VScroll1_Change()
Dim i As Integer
For i = VScroll1.Value To 99
 txtIndex(i).Top = txtIndex(i - VScroll1.Value).Top + 390
 txtNumber(i).Top = txtNumber(i - VScroll1.Value).Top + 390
 txtName(i).Top = txtNumber(i - VScroll1.Value).Top + 390
 txtBan(i).Top = txtBan(i - VScroll1.Value).Top + 390
 txtChe(i).Top = txtChe(i - VScroll1.Value).Top + 390
 txtFlag(i).Top = txtFlag(i - VScroll1.Value).Top + 390
 txtnow_che(i).Top = txtnow_che(i - VScroll1.Value).Top + 390
Next
End Sub



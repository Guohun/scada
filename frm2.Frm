VERSION 5.00
Begin VB.Form zu1 
   Caption         =   "Printer Options"
   ClientHeight    =   1776
   ClientLeft      =   1092
   ClientTop       =   1488
   ClientWidth     =   6768
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1776
   ScaleWidth      =   6768
   Begin VB.CommandButton Command2 
      Caption         =   "Reset Number of Copies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept Number of Copies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Num_Copies 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Text            =   "1"
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Please enter the number of copies you wish to print :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "zu1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Number_Of_Copies = Num_Copies
    Form2.Hide
   
End Sub

Private Sub Command2_Click()
    Num_Copies.SetFocus
    Num_Copies.Text = "1"
End Sub


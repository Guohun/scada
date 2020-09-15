VERSION 5.00
Begin VB.Form xiao_liao_table 
   Caption         =   "Form1"
   ClientHeight    =   4164
   ClientLeft      =   1740
   ClientTop       =   1932
   ClientWidth     =   6972
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4164
   ScaleWidth      =   6972
   Begin VB.Frame Frame2 
      Caption         =   "Move Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   6612
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label record_prompt 
         Caption         =   "Label1"
         Height          =   372
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   1812
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   6735
      Begin VB.TextBox Text4 
         Height          =   372
         Left            =   4920
         TabIndex        =   20
         Text            =   "Text4"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox Text3 
         Height          =   372
         Left            =   4920
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox Text2 
         Height          =   408
         Left            =   1800
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Left            =   1800
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   372
         Left            =   3240
         TabIndex        =   19
         Top             =   1080
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   372
         Left            =   3240
         TabIndex        =   17
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   372
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   372
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Width           =   1212
      End
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H000000C0&
      Caption         =   "&Add  Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H000000C0&
      Caption         =   "&Edit Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000C0&
      Caption         =   "&Del Record"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      TabIndex        =   3
      Top             =   0
      Width           =   1332
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5280
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Data Dxiao_liao_table 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\ljxt\ljxt.Mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "xiao_liao_table"
      Top             =   3360
      Visible         =   0   'False
      Width           =   6732
   End
   Begin VB.CommandButton exit_command 
      Caption         =   "Command1"
      Height          =   372
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   1092
   End
End
Attribute VB_Name = "xiao_liao_table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

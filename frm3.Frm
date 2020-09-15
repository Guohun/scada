VERSION 5.00
Begin VB.Form zu2 
   Caption         =   "Print To File Settings"
   ClientHeight    =   2880
   ClientLeft      =   1092
   ClientTop       =   1488
   ClientWidth     =   8076
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   8076
   Begin VB.CommandButton Command2 
      Caption         =   "Modify Settings"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept Settings"
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
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
   Begin VB.ComboBox OutputFileType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Text            =   "  "
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox OutputFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Text            =   "c:\crw\OutFile.txt"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Select the output type :"
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
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter output file name  ( including path ) :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "zu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Output_File_Name = OutputFileName.Text
   
   Rem - Set PrinterFileType property
    If OutputFileType = "Record Format" Then
        Output_File_Type = 0
    Else
        If OutputFileType = "Tab Separated" Then
            Output_File_Type = 1
        Else
            If OutputFileType = "Text Format" Then
                Output_File_Type = 2
            Else
                If OutputFileType = "DIF Format" Then
                    Output_File_Type = 3
                Else
                    If OutputFileType = "Comma Separated Value" Then
                        Output_File_Type = 4
                    Else
                        If OutputFileType = "Tab Separated Text" Then
                            Output_File_Type = 6
                        End If
                    End If
                End If
            End If
        End If
    End If

Rem Close form
'    Form3.Hide

End Sub

Private Sub Command2_Click()
    OutputFileType.SetFocus
End Sub

Private Sub Form_Load()
    OutputFileType.AddItem "Comma Separated"
    OutputFileType.AddItem "Tab Separated"
    OutputFileType.AddItem "Record Format"
    OutputFileType.AddItem "Text Format"
    OutputFileType.AddItem "DIF Format"
    OutputFileType.AddItem "Tab Separated Text"
    OutputFileType.SelText = "Comma Separated"

    
End Sub


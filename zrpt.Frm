VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   195
   ClientTop       =   2025
   ClientWidth     =   6690
   Height          =   4170
   Left            =   135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6690
   Top             =   1680
   Width           =   6810
   Begin VB.CommandButton cmd1 
      Caption         =   "打开报表"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1990
   End
   Begin VB.TextBox txt 
      Height          =   495
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "设份数打印"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1440
      TabIndex        =   1
      Top             =   3000
      Width           =   1990
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "打印至屏幕"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   1990
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   6240
      Top             =   120
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ReportFileName  =   ""
      Destination     =   0
      WindowLeft      =   100
      WindowTop       =   100
      WindowWidth     =   490
      WindowHeight    =   300
      WindowTitle     =   ""
      WindowBorderStyle=   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      CopiesToPrinter =   1
      PrintFileName   =   ""
      PrintFileType   =   0
      SelectionFormula=   ""
      GroupSelectionFormula=   ""
      Connect         =   ""
      UserName        =   ""
      ReportSource    =   0
      BoundReportHeading=   ""
      BoundReportFooter=   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
   Dim A As Integer
Private Sub cmd1_Click()
      cdg.Filter = "报表文件类型(*.rpt)*.rpt"
      cdg.Action = 1
      rpt.ReportFileName = cdg.filename
      cmd2.Enabled = True
      cmd3.Enabled = True
      rpt.WindowTitle = cdg.filename
      
End Sub

Private Sub cmd2_Click()
       rpt.Destination = 0
       rpt.Action = 1
End Sub

Private Sub cmd3_Click()
          
         A = Val(txt.Text)
          
          rpt.Destination = 1
          rpt.CopiesToPrinter = A
        '  0n error goto errorhandler
          
          rpt.Action = 1
      '  Exit Sub
        
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
     
       txt.Text = 1
       A = txt.Text
    
    
    
    cmd2.Enabled = False
    cmd3.Enabled = False
End Sub

Private Sub txt_Change()
    
    '   txt.Text = 1
     '  A = Val(txt.Text)
End Sub

Private Sub Text1_Change()

End Sub

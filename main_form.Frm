VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.MDIForm MAIN_MDI 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6876
   ClientLeft      =   -852
   ClientTop       =   -1.38492e5
   ClientWidth     =   9480
   Icon            =   "main_form.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Begin VB.PictureBox Button_Picture 
      Align           =   1  'Align Top
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12
      Left            =   0
      ScaleHeight     =   12
      ScaleWidth      =   9480
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      Begin VB.Timer Timer3 
         Interval        =   800
         Left            =   2280
         Top             =   96
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   1800
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   1320
         Top             =   120
      End
      Begin VB.CommandButton cmdenglish 
         Caption         =   "Enlish Version"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdchina 
         Caption         =   "中文版"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Buttom_Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   7956
         TabIndex        =   1
         Top             =   4800
         Width           =   1176
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt_Label"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   42
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   840
         Left            =   2184
         TabIndex        =   5
         Top             =   2052
         Width           =   5208
      End
      Begin VB.Label Prompt_Label 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt_Label"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   42
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2240
         TabIndex        =   2
         Top             =   2120
         Width           =   5256
      End
   End
   Begin Threed.SSPanel StatuBar 
      Align           =   2  'Align Bottom
      Height          =   456
      Left            =   0
      TabIndex        =   6
      Top             =   6420
      Width           =   9480
      _Version        =   65536
      _ExtentX        =   16722
      _ExtentY        =   804
      _StockProps     =   15
      BackColor       =   8388736
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin ComctlLib.ProgressBar ProcessBar 
         Height          =   252
         Left            =   7200
         TabIndex        =   9
         Top             =   96
         Visible         =   0   'False
         Width           =   2172
         _ExtentX        =   3831
         _ExtentY        =   445
         _Version        =   327680
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label Process_Msg 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5640
         TabIndex        =   10
         Top             =   -1440
         Visible         =   0   'False
         Width           =   1512
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "运行状态监视："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   96
         Width           =   1812
      End
      Begin VB.Label Run_Status 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   13.8
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   252
         Left            =   3000
         TabIndex        =   7
         Top             =   96
         Width           =   3660
      End
   End
   Begin VB.Menu mnuad1 
      Caption         =   "ad1"
      Visible         =   0   'False
      Begin VB.Menu mnuada 
         Caption         =   "ada"
         Index           =   0
      End
   End
   Begin VB.Menu mnuad2 
      Caption         =   "ad2"
      Visible         =   0   'False
      Begin VB.Menu mnuadb 
         Caption         =   "adb"
         Index           =   0
      End
   End
   Begin VB.Menu mnuad3 
      Caption         =   "ad3"
      Visible         =   0   'False
      Begin VB.Menu mnuadc 
         Caption         =   "adc"
         Index           =   0
      End
   End
   Begin VB.Menu mnuad4 
      Caption         =   "ad4"
      Visible         =   0   'False
      Begin VB.Menu mnuadd 
         Caption         =   "add"
         Index           =   0
      End
   End
   Begin VB.Menu mnuad5 
      Caption         =   "ad5"
      Visible         =   0   'False
      Begin VB.Menu mnuade 
         Caption         =   "ade"
         Index           =   0
      End
   End
   Begin VB.Menu mnuad6 
      Caption         =   "ad6"
      Visible         =   0   'False
      Begin VB.Menu mnuadf 
         Caption         =   "adf"
         Index           =   0
      End
   End
   Begin VB.Menu mnuad7 
      Caption         =   "ad7"
      Visible         =   0   'False
      Begin VB.Menu mnuadg 
         Caption         =   "adg"
         Index           =   0
      End
   End
   Begin VB.Menu mnuad8 
      Caption         =   "ad8"
      Visible         =   0   'False
      Begin VB.Menu mnuadh 
         Caption         =   "adh"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MAIN_MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numbera As Integer
Dim numberb As Integer
Dim numberc As Integer
Dim numberd As Integer
Dim numbere As Integer
Dim numberf As Integer
Dim numberg As Integer


Dim gname(50) As String
Dim gfather(50) As String
Dim gnext(50) As String
Dim gprev(50) As String
Dim grow(50) As Integer
Dim gcol(50) As Integer
Dim gchina(50) As String
Dim genglish(50) As String
Dim gkey(50) As String
Dim gflag(50) As String
Dim gbrothermenu(50) As String
Dim gbrothertemp(50) As String

Dim gnames(50) As String
Dim gfathers(50) As String
Dim gnexts(50) As String
Dim gprevs(50) As String
Dim grows(50) As Integer
Dim gcols(50) As Integer
Dim gchinas(50) As String
Dim genglishs(50) As String
Dim gkeys(20) As String
Dim Sub_Keys(20, 20) As String
Dim gflags(50) As String
Dim gbrothermenus(50) As String
Dim gbrothertemps(50) As String

Dim glastadd As Integer

Dim glastada As Integer
Dim glastadc As Integer
Dim glastadb As Integer
Dim glastade As Integer
Dim glastadf As Integer
Dim glastadg As Integer
Dim gcount As Integer
Dim gfathercount As Integer
Dim gsuncount As Integer
Dim mnusavea(50) As String
Dim mnusaveb(50) As String
Dim mnusavec(50) As String
Dim mnusaved(50) As String
Dim mnusavee(50) As String
Dim mnusavef(50) As String
Dim mnusaveg(50) As String

Private Sub Button_Picture_KeyUp(KeyCode As Integer, Shift As Integer)
        Call Check_KeyCode(KeyCode, Shift)
End Sub

'************************
'中文按钮事件
'************************
Private Sub cmdchina_Click()
Dim i As Integer
Dim j As Integer
Dim k As Integer '
Dim numbera As Integer
Dim numberb As Integer
Dim numberc As Integer
Dim numberd As Integer
Dim numbere As Integer
Dim numberf As Integer
Dim numberg As Integer
numbera = 0
numberb = 0
numberc = 0
numberd = 0
numbere = 0
numberf = 0
numberg = 0


' 放置父菜单
For i = 0 To gfathercount - 1
 Select Case i
 Case 0:
        mnuad1.Visible = 1
       mnuad1.Caption = "&" & gkey(i) & gchina(i)
Case 1:
        mnuad2.Visible = 1
       mnuad2.Caption = "&" & gkey(i) & gchina(i)
Case 2:
       mnuad3.Visible = 1
       mnuad3.Caption = "&" & gkey(i) & gchina(i)
Case 3:
       mnuad4.Visible = 1
       mnuad4.Caption = "&" & gkey(i) & gchina(i)
Case 4:
       mnuad5.Visible = 1
       mnuad5.Caption = "&" & gkey(i) & gchina(i)
Case 5:
       mnuad6.Visible = 1
       mnuad6.Caption = "&" & gkey(i) & gchina(i)
Case 6
       mnuad7.Visible = 1
       mnuad7.Caption = "&" & gkey(i) & gchina(i)
End Select
Next i
' 安置总子菜单
'For j = 0 To gfathercount - 1
 For k = 0 To gsuncount - 1
   If gfathers(k) = gchina(0) Then
       If gcols(k) <> 0 Then
           mnuada(gcols(k) - 1).Caption = "&" & Sub_Keys(0, gcols(k)) & gchinas(k)
       End If
     End If
   If gfathers(k) = gchina(1) Then
          If gcols(k) <> 0 Then
           mnuadb(gcols(k) - 1).Caption = "&" & Sub_Keys(1, gcols(k)) & gchinas(k)
          End If
      End If
   If gfathers(k) = gchina(2) Then
       If gcols(k) <> 0 Then
           mnuadc(gcols(k) - 1).Caption = "&" & Sub_Keys(2, gcols(k)) & gchinas(k)
       End If
     End If
   If gfathers(k) = gchina(3) Then
       If gcols(k) <> 0 Then
           mnuadd(gcols(k) - 1).Caption = "&" & Sub_Keys(3, gcols(k)) & gchinas(k)
        End If
      End If
   If gfathers(k) = gchina(4) Then
       If gcols(k) <> 0 Then
         mnuade(gcols(k) - 1).Caption = "&" & Sub_Keys(4, gcols(k)) & gchinas(k)
       End If
      End If
   If gfathers(k) = gchina(5) Then
    If gcols(k) <> 0 Then
        mnuadf(gcols(k) - 1).Caption = "&" & Sub_Keys(5, gcols(k)) & gchinas(k)
    End If
     End If
   If gfathers(k) = gchina(6) Then
     If gcols(k) <> 0 Then
       mnuadg(gcols(k) - 1).Caption = "&" & Sub_Keys(6, gcols(k)) & gchinas(k)
     End If
      End If
   Next k
 'Next j
 '计算各父菜单无用子菜单数目
  For k = 0 To gsuncount - 1
     If gfathers(k) = gchina(0) Then
        numbera = numbera + 1
      End If
      If gfathers(k) = gchina(1) Then
        numberb = numberb + 1
      End If
     If gfathers(k) = gchina(2) Then
        numberc = numberc + 1
      End If
      If gfathers(k) = gchina(3) Then
        numberd = numberd + 1
      End If
     If gfathers(k) = gchina(4) Then
        numbere = numbere + 1
      End If
     If gfathers(k) = gchina(5) Then
        numberf = numberf + 1
      End If
      If gfathers(k) = gchina(6) Then
        numberg = numberg + 1
      End If
    Next k
  i = numbera
  For k = numbera To gsuncount - 1
     If mnuada(k).Caption = "ada" Then
       mnuada(k).Visible = 0
      End If
   Next k
 '清除无用子菜单
   For k = numberb To gsuncount - 1
     If mnuadb(k).Caption = "adb" Then
       mnuadb(k).Visible = 0
      End If
   Next k
  For k = numberc To gsuncount - 1
     If mnuadc(k).Caption = "adc" Then
       mnuadc(k).Visible = 0
      End If
   Next k
  If numberd = 0 Then
    numberd = 1
    End If
  For k = numberd To gsuncount - 1
     If mnuadd(k).Caption = "add" Then
       mnuadd(k).Visible = 0
      End If
    Next k
  For k = numbere To gsuncount - 1
     If mnuade(k).Caption = "ade" Then
       mnuade(k).Visible = 0
      End If
   Next k
 If numberf = 0 Then
    numberf = 1
    End If
    
   For k = numberf To gsuncount - 1
     If mnuadf(k).Caption = "adf" Then
       mnuadf(k).Visible = 0
      End If
     Next k
 If numberg = 0 Then
    numberg = 1
    End If
      
   For k = numberg To gsuncount - 1
     If mnuadg(k).Caption = "adg" Then
       mnuadg(k).Visible = 0
      End If
  Next k
  
End Sub

'************************
'英文按钮事件
'************************
Private Sub cmdenglish_Click()
Dim i As Integer
Dim j As Integer
Dim k As Integer '
Dim numbera As Integer
Dim numberb As Integer
Dim numberc As Integer
Dim numberd As Integer
Dim numbere As Integer
Dim numberf As Integer
Dim numberg As Integer
numbera = 1
numberb = 1
numberc = 1
numberd = 1
numbere = 1
numberf = 1
numberg = 1


' 放置父菜单
For i = 0 To gfathercount - 1
 Select Case i
 Case 0:
    mnuad1.Visible = 1
   mnuad1.Caption = "&" & gkey(i) & genglish(i)
Case 1:
    mnuad2.Visible = 1
   mnuad2.Caption = "&" & gkey(i) & genglish(i)
Case 2:
   mnuad3.Visible = 1
   mnuad3.Caption = "&" & gkey(i) & genglish(i)
Case 3:
   mnuad4.Visible = 1
   mnuad4.Caption = "&" & gkey(i) & genglish(i)
Case 4:
   mnuad5.Visible = 1
   mnuad5.Caption = "&" & gkey(i) & genglish(i)
Case 5:
   mnuad6.Visible = 1
   mnuad6.Caption = "&" & gkey(i) & genglish(i)
Case 6
   mnuad7.Visible = 1
   mnuad7.Caption = "&" & gkey(i) & genglish(i)
End Select
Next i
' 安置总子菜单
 For k = 0 To gsuncount - 1
   If gfathers(k) = gchina(0) Then
       mnuada(gcols(k) - 1).Caption = "&" & Sub_Keys(0, gcols(k)) & genglishs(k)
     End If
   If gfathers(k) = gchina(1) Then
       mnuadb(gcols(k) - 1).Caption = "&" & Sub_Keys(1, gcols(k)) & genglishs(k)
      End If
   If gfathers(k) = gchina(2) Then
       mnuadc(gcols(k) - 1).Caption = "&" & Sub_Keys(2, gcols(k)) & genglishs(k)
     End If
   If gfathers(k) = gchina(3) Then
       mnuadd(gcols(k) - 1).Caption = "&" & Sub_Keys(3, gcols(k)) & genglishs(k)
      End If
   If gfathers(k) = gchina(4) Then
       mnuade(gcols(k) - 1).Caption = "&" & Sub_Keys(4, gcols(k)) & genglishs(k)
      End If
   If gfathers(k) = gchina(5) Then
       mnuadf(gcols(k) - 1).Caption = "&" & Sub_Keys(5, gcols(k)) & genglishs(k)
     End If
   If gfathers(k) = gchina(6) Then
       mnuadg(gcols(k) - 1).Caption = "&" & Sub_Keys(6, gcols(k)) & genglishs(k)
      End If
   Next k
   
 '计算各父菜单无用子菜单数目
  For k = 0 To gsuncount - 1
     If gfathers(k) = genglish(0) Then
        numbera = numbera + 1
      End If
      If gfathers(k) = genglish(1) Then
        numberb = numberb + 1
      End If
     If gfathers(k) = genglish(2) Then
        numberc = numberc + 1
      End If
      If gfathers(k) = genglish(3) Then
        numberd = numberd + 1
      End If
     If gfathers(k) = genglish(4) Then
        numbere = numbere + 1
      End If
     If gfathers(k) = genglish(5) Then
        numberf = numberf + 1
      End If
      If gfathers(k) = genglish(6) Then
        numberg = numberg + 1
      End If
    Next k


  i = numbera
  For k = numbera To gsuncount - 1
     If mnuada(k).Caption = "ada" Then
       mnuada(k).Visible = 0
      End If
   Next k
 '清除无用子菜单
   For k = numberb To gsuncount - 1
     If mnuadb(k).Caption = "adb" Then
       mnuadb(k).Visible = 0
      End If
   Next k
  For k = numberc To gsuncount - 1
     If mnuadc(k).Caption = "adc" Then
       mnuadc(k).Visible = 0
      End If
   Next k
  If numberd = 0 Then
    numberd = 1
    End If
  For k = numberd To gsuncount - 1
     If mnuadd(k).Caption = "add" Then
       mnuadd(k).Visible = 0
      End If
    Next k
  For k = numbere To gsuncount - 1
     If mnuade(k).Caption = "ade" Then
       mnuade(k).Visible = 0
      End If
   Next k
 If numberf = 0 Then
    numberf = 1
    End If
    
   For k = numberf To gsuncount - 1
     If mnuadf(k).Caption = "adf" Then
       mnuadf(k).Visible = 0
      End If
     Next k
 If numberg = 0 Then
    numberg = 1
    End If
      
   For k = numberg To gsuncount - 1
     If mnuadg(k).Caption = "adg" Then
       mnuadg(k).Visible = 0
      End If
  Next k
  

End Sub



Private Sub MDIForm_Activate()
    If MAIN_MDI!Button_Picture.Visible = True Then
        SendKeys "%1", True  ' Send ALT+F4 to close Calculator.
    End If
End Sub




Private Sub MDIForm_Load()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim mnudata As Database
Dim recset As Recordset
Width = Screen.Width '* 0.75 ' Set width of form.
Call Read_System
    If Multi_Screen = "2" Then
            Height = Screen.Height * 0.5  ' Set height of form.
    Else
            Height = Screen.Height ' * 0.75   ' Set height of form.
    End If
    Left = 0
    Top = 0

Call First_Do
Call Init_P
Call Init_Elec
Run_Hwnd = ExecProgram(MAIN_MDI.Hwnd, "c:\ljxt\ljm.exe")
Get_Focus_Fore_Color = RGB(255, 0, 0)
Los_Focus_Fore_Color = &HC00000

glastada = 0
glastadc = 0
glastadb = 0
glastadd = 0
glastade = 0
glastadf = 0
glastadg = 0
Set mnudata = OpenDatabase("c:\ljxt\ljxt.mdb")
Set recset = mnudata.OpenRecordset("menu_table", dbOpenDynaset)


recset.MoveLast
gcount = recset.RecordCount
recset.MoveFirst
k = 0
For i = 0 To gcount - 1
    If recset.Fields("father_name").Value = "0" Then
        gname(k) = recset.Fields("name").Value
        gfather(k) = recset.Fields("father_name").Value
        gnext(k) = recset.Fields("next_name").Value
        gprev(k) = recset.Fields("prev_name").Value
        grow(k) = recset.Fields("row").Value
        gcol(k) = recset.Fields("col").Value
        gchina(k) = recset.Fields("china_name").Value
        genglish(k) = recset.Fields("english_name").Value
        gkey(k) = recset.Fields("key").Value
        gflag(k) = recset.Fields("flag").Value
        k = k + 1
        gfathercount = gfathercount + 1
    End If
    recset.MoveNext
Next i
recset.MoveFirst
k = 0
For j = 0 To gcount - 1
If recset.Fields("father_name").Value <> "0" Then
            gnames(k) = vFieldVal(recset.Fields("name").Value)
            gfathers(k) = recset.Fields("father_name").Value
            gnexts(k) = vFieldVal(recset.Fields("next_name").Value)
            gprevs(k) = vFieldVal(recset.Fields("prev_name").Value)
            grows(k) = recset.Fields("row").Value
            gcols(k) = recset.Fields("col").Value
            gchinas(k) = recset.Fields("china_name").Value
            genglishs(k) = recset.Fields("english_name").Value
            Sub_Keys(grows(k), gcols(k)) = recset.Fields("key").Value       '行列
            'gbrothermenus(k) = recset.Fields("next_menu").Value
            'gbrothertemps(k) = recset.Fields("next_temp").Value
            gsuncount = gsuncount + 1
        k = k + 1
        End If
        recset.MoveNext
Next j
recset.MoveFirst
'MsgBox "The total, father and sub menus number is:" + Str(gcount) + Str(gfathercount) + Str(gsuncount)
For k = 1 To gsuncount - 1
    Load mnuadb(k)
    Load mnuada(k)
    Load mnuadc(k)
    Load mnuadd(k)
    Load mnuade(k)
    Load mnuadf(k)
    Load mnuadg(k)
    mnusavea(k) = mnuada(k)
    mnusaveb(k) = mnuadb(k)
    mnusavec(k) = mnuadc(k)
    mnusaved(k) = mnuadd(k)
    mnusavee(k) = mnuade(k)
    mnusavef(k) = mnuadf(k)
    mnusaveg(k) = mnuadg(k)
Next k

Call cmdchina_Click


End Sub



Private Sub First_Do()
     If China_English = CHINA Then
             Me.Caption = "  炼胶系统 "
             Buttom_Label.Caption = "化学工业部桂林橡胶工业设计院"
             Prompt_Label.Caption = "炼胶生产自动控制系统"
             
    Else
            Prompt_Label.Caption = "LJ  Auto  Produce  Control  System"
            Buttom_Label.Caption = "Guilin Design &Research Institute of Rubber Industry Minstry of Chemical  Industry"
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        Cancel = 1
End Sub

Private Sub MDIForm_Resize()
  If WindowState <> 1 Then   'not minimized
  If MAIN_MDI.Height - MAIN_MDI.Top - 700 > 0 Then
      Button_Picture.Height = MAIN_MDI.Height - MAIN_MDI.Top - 700
 End If
   End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    Cancel = 1
End Sub



'*********************************
'菜单a命令
'*********************************
Private Sub mnuada_Click(index As Integer)
Dim answer   As Integer
Select Case index     'Mid(mnuada(index).Caption, 3)
  Case 2     'Is = "退出", "Exit"
      If Make_Mathine = 0 Then
            Select Case Fen_Run_flag
                    Case 0, 10
                    If Run_Hwnd <> 0 Then
                            Call TerminateProcess(Run_Hwnd, 0)
                    End If
                    Unload Me
                    End
                    Case 1, 2, 3
                    answer = MsgBox(" 在运行  确定退出吗？退出时要等待", 48 + vbYesNo, "提示")
                    If answer = vbYes Then
                        Call TerminateProcess(Run_Hwnd, 0)
                        Fen_Run_flag = 0
                        Unload Me
                        Call ShowCursor(1)
                        End
                    End If
                    Case 11
                    answer = MsgBox(" 在模拟  确定退出吗？退出时要等待", 48 + vbYesNo, "提示")
                    If answer = vbYes Then
                        If Simulate_Hwnd <> 0 Then
                            Call TerminateProcess(Simulate_Hwnd, 0)
                        End If
                        Run_flag = 0
                        Unload Me
                        Call ShowCursor(1)
                    End
                    End If
                    
            End Select
      End If
    Select Case Run_flag
    Case 0
            If Run_Hwnd <> 0 Then
                Call TerminateProcess(Run_Hwnd, 0)
            End If
            Unload Me
             End
    Case 1
        answer = MsgBox(" 在运行  确定退出吗？退出时要等待", 48 + vbYesNo, "提示")
       If answer = vbYes Then
            Call TerminateProcess(Run_Hwnd, 0)
            Run_flag = 0
            Unload Me
            Call ShowCursor(1)
            End
    End If
    Case 2
        answer = MsgBox(" 在模拟  确定退出吗？退出时要等待", 48 + vbYesNo, "提示")
        If answer = vbYes Then
            If Simulate_Hwnd <> 0 Then
                    Call TerminateProcess(Simulate_Hwnd, 0)
            End If
            Run_flag = 0
            Unload Me
            Call ShowCursor(1)
            End
        End If
    Case 3
        answer = MsgBox(" 在线修改  确定退出吗？退出时要等待", 48 + vbYesNo, "提示")
        If answer = vbYes Then
            If Run_Hwnd <> 0 Then
                Call TerminateProcess(Run_Hwnd, 0)
            End If
            If Simulate_Hwnd <> 0 Then
                    Call TerminateProcess(Simulate_Hwnd, 0)
            End If
            
            Run_flag = 0
            Unload Me
            Call ShowCursor(1)
            End
        End If
    
    End Select
  Case 0  'Is = "生产命令录入", "Produce"
                produce.Show
                produce.ZOrder 0
   Case 1 'Is = "风送系统", "Fen_Send"
                frmFen_Pei_table.Show
                frmFen_Pei_table.ZOrder 0
End Select
End Sub

'*********************************
'菜单b命令
'*********************************

Private Sub mnuadb_Click(index As Integer)

Select Case index
  Case 0
                   Prep_Cai_Liao.Show
                   Prep_Cai_Liao.ZOrder 0
  Case 1
                   Pei_fan_table_Input.Show
                   Pei_fan_table_Input.ZOrder 0
End Select

End Sub

'*********************************
'菜单c命令
'*********************************
Private Sub mnuadc_Click(index As Integer)
Select Case index
          Case 0
                    Elec_Run.Show
                    Elec_Run.ZOrder 0
                        
          Case 1
                        
                         Elec_Set.Show
                         Elec_Set.ZOrder 0
  End Select
               

End Sub

'*********************************
'菜单d命令
'*********************************
Private Sub mnuadd_Click(index As Integer)
Select Case index
          Case 4
                        zhu_pro.Tag = "tempro"
                        zhu_pro.Show
          Case 5
                        zhu_pro.Show
                        zhu_pro.Tag = "power"
          Case 1
                        zhu.Show
                        zhu.Tag = 1
          Case 2
                        zhu.Show
                        zhu.Tag = 2
          Case 3
                        zhu_pro.Show
                        zhu_pro.Tag = "produce"
          Case 0
                        With F_ShowLc
                            .Enabled = True
                            .Timer1.Enabled = True
                            .Timer2.Enabled = True
                            .Show
                            .SetFocus
                        End With
            Case 6
                    If Password_Do("Del_Produce", Mid(mnuadd(index).Caption, 3)) = False Then
                                Exit Sub
                    End If
                    Dim Message, Title, Default, MyValue
                    Message = "      请输入指定日期，该数据包含指定日期以前的温度，能量图形数据" _
                                    & " 生产数据    默认日期为  系统一个月以前的时间"
                    Title = "删除窗口" ' Set title.
                    Default = DateAdd("m", -1, Date$) ' Set default.
                
                    MyValue = InputBox(Message, Title, Default)
                    If IsDate(MyValue) Then
                        Del_Make_Data (MyValue)
                    Else
                        Speak_Error ("非法日期: " & MyValue)
                    End If
                     MAIN_MDI.SetFocus
                     SendKeys "%4", True
   End Select

End Sub

'*********************************
'菜单e命令
'*********************************
Private Sub mnuade_Click(index As Integer)
Select Case index
  Case 0
         China_English = Not China_English
         If China_English = CHINA Then
            Call cmdchina_Click
            Call Change_Language
         Else
            Call cmdenglish_Click
            Call Change_Language
         End If
         Call First_Do
  Case 1
                Menu_table_Input.Show
  Case 4
                     comm_set_Form.Show
                     
                     
  Case 2
                    send_table_input.Show
  Case 3
                     Fen_send_table.Show
                     
  Case 5
                    send_file_form.Show
  Case 6
             Date_form.Show
  End Select
End Sub
'*********************************
'菜单f命令
'*********************************
Private Sub mnuadf_Click(index As Integer)
Select Case index
  Case 0
         error_Input.Show
   Case 1
            Elec_not_table_Input.Show
End Select

End Sub

'*********************************
'菜单g命令
'*********************************

Private Sub mnuadg_Click(index As Integer)
Select Case Mid(mnuadg(index).Caption, 3)
  Case Is = "图象蛇":
         Error.Show
   Case Is = "故障测量表":
            Elec_not_table_Input.Show
End Select

End Sub

Private Sub Prompt_Label_Change()
    Label1.Caption = Prompt_Label.Caption
End Sub




Private Sub Timer1_Timer()
    Static X
    X = X + 1
    If X > 10 Then
        Buttom_Label.Left = Buttom_Label.Left - 100
        If (Buttom_Label.Left + Buttom_Label.Width) <= 0 Then
            Buttom_Label.Left = MAIN_MDI.Width
            
        End If
        X = 0
    End If
    If Make_Mathine = 0 Then
        Call read_Fs_command(Fen_Run_flag)
            Select Case Fen_Run_flag
                Case 1
                    
                    If Prev_Fen_Run_flag = 2 Then
                        Run_Status.Caption = "在线修改" + space(4) + Trim(Run_Do_Table.number) + "配方名"
                    End If
                Case 2
                    Run_Status.Caption = "运行" + space(4) + Trim(Run_Do_Table.number) + "配方名"
                    Prev_Fen_Run_flag = Fen_Run_flag
                Case 3
                    Run_Status.Caption = "结束" + space(4) + Trim(Run_Do_Table.number) + "配方名"
                Case 4
                    Run_Status.Caption = "修改确认" + space(4) + Trim(Run_Do_Table.number) + "配方名"
                Case 11
                    Run_Status.Caption = "风送系统模拟" + space(4) + Trim(Run_Do_Table.number) + "配方名"
                Case 10
                    Run_Status.Caption = "停止"
                Case Else
                    Run_Status.Caption = "停止"
            End Select
        Exit Sub
    End If
    Call read_run(Run_Do_Table, Run_flag)
    If China_English = CHINA Then
        Select Case Run_flag
            Case 1
                Run_Status.Caption = "运行" + space(4) + Trim(Run_Do_Table.number) + "配方名"
            Case 2
                Run_Status.Caption = "模拟" + space(4) + Trim(Run_Do_Table.number) + "配方名"
            Case 3
                Run_Status.Caption = "在线修改" + space(4) + Trim(Run_Do_Table.number) + "配方名"
            Case 4
                Run_Status.Caption = "修改确认" + space(4) + Trim(Run_Do_Table.number) + "配方名"
            Case 10
                Run_Status.Caption = "停止"
            Case Else
                Run_Status.Caption = "停止"
        End Select
    End If
End Sub


Private Sub Timer2_Timer()
    Dim temp_number   As String
    Dim temp_name   As String
    Dim temp_ban     As Integer
    Dim temp_che    As Integer
    Dim temp_flag     As Integer
    Dim Mathine%
    Dim Z As Long
    Dim DbBase  As Database
    Dim DbRec  As Recordset
    On Error Resume Next
    If Run_flag = 10 Then
        Timer2.Enabled = False
        Run_flag = 0
        GoTo simulate_exit
    End If
    If Run_flag = 0 Then
            Set DbBase = OpenDatabase("c:\ljxt\ljxt.mdb", , True)
            Make_Row = Make_Row + 1
    
            Set DbRec = DbBase.OpenRecordset("select * from run_table where index>=" & Make_Row & "  order by index", dbOpenDynaset, dbReadOnly)
            If DbRec.EOF Then
                       Run_flag = 0
                        Call Speak_Error("所有配方运行结束  ")
                       GoTo simulate_exit
            End If
            
            If DbRec.Fields("index") = "" Then
                Call Speak_Error("第" & Make_Row & "记录的" & " 目录没有输入")
                GoTo simulate_exit
            End If
            If DbRec.Fields("_number") = "" Or IsEmpty(DbRec.Fields("_number")) Then
                Call Speak_Error("第" & Make_Row & "记录的" & " 配方编号没有输入")
                GoTo simulate_exit
            End If
            If DbRec.Fields("ban") = "" Or IsNull(DbRec.Fields("ban")) Then
                Call Speak_Error("第" & Make_Row & "记录的" & " 班号没有输入")
                GoTo simulate_exit
            End If
            If DbRec.Fields("che") = "" Or IsNull(DbRec.Fields("che")) Then
                Call Speak_Error("第" & Make_Row & "记录的" & " 车数没有输入")
                GoTo simulate_exit
            End If
            Make_Row = vFieldVal(DbRec.Fields("index"))
            temp_number = vFieldVal(DbRec.Fields("_number"))
            temp_name = vFieldVal(DbRec.Fields("name"))
            temp_ban = Val(DbRec.Fields("ban"))
            temp_che = Val(DbRec.Fields("che"))
            Mathine = Val(DbRec.Fields("mathine"))
            
            Screen.MousePointer = vbDefault
            produce!Simulate_Command.Enabled = False
            produce!cmdRefresh.Enabled = False
            produce!Change_Command.Enabled = True
            If Make_Row = 1 Then
                F_ShowLc.Show
                SendKeys "%4", True  ' Send UpKey to close Calculator.
                SendKeys "1", True  ' Send UpKey to close Calculator.
             End If
            Call Sleep(100)
            If Prev_run_flag = False Then
'                                Run_Hwnd = ExecProgram(MAIN_MDI.Hwnd, "c:\ljxt\ljm.exe")
                                Prev_run_flag = True
            End If
            If can_read(temp_number, temp_name, temp_ban, temp_che, 16 Or Mathine) = 0 Then
                            Call Speak_Error("没设置ODBC 或ljxt.mdb数据库有错误")
                            GoTo simulate_exit
            End If
            Run_flag = 1
            DbRec.Close
            DbBase.Close
    End If
    Exit Sub
simulate_exit:
            Make_Row = 0
            Timer2.Enabled = False
            DbRec.Close
            DbBase.Close
            produce!Simulate_Command.Enabled = True
            produce!cmdRefresh.Enabled = True
            produce!Change_Command.Enabled = False
End Sub

Private Sub Timer3_Timer()
    If Multi_Screen <> "0" Then
          Call VBwrite_multi_screen
    End If
End Sub

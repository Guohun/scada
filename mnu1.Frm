VERSION 4.00
Begin VB.Form frmmenu 
   Caption         =   " the program of menu"
   ClientHeight    =   4260
   ClientLeft      =   525
   ClientTop       =   1470
   ClientWidth     =   7560
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   4665
   Left            =   465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   7560
   Top             =   1125
   Width           =   7680
   Begin VB.CommandButton cmdenglish 
      Caption         =   "Enlish Version"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdchina 
      Caption         =   "中文版"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
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
End
Attribute VB_Name = "frmmenu"
Attribute VB_Creatable = False
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
Dim gkeys(50) As String
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
Dim mnudata As Database
Dim recset As Recordset


'Dim resource As RecordSource




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
       mnuad1.Caption = gchina(i)
Case 1:
        mnuad2.Visible = 1
       mnuad2.Caption = gchina(i)
Case 2:
       mnuad3.Visible = 1
       mnuad3.Caption = gchina(i)
Case 3:
       mnuad4.Visible = 1
       mnuad4.Caption = gchina(i)
Case 4:
       mnuad5.Visible = 1
       mnuad5.Caption = gchina(i)
Case 5:
       mnuad6.Visible = 1
       mnuad6.Caption = gchina(i)
Case 6
       mnuad7.Visible = 1
       mnuad7.Visible = gchina(i)
End Select
Next i
' 安置总子菜单
'For j = 0 To gfathercount - 1
 For k = 0 To gsuncount - 1
   If gfathers(k) = gchina(0) Then
       If gcols(k) <> 0 Then
           mnuada(gcols(k) - 1).Caption = gchinas(k)
       End If
     End If
   If gfathers(k) = gchina(1) Then
          If gcols(k) <> 0 Then
           mnuadb(gcols(k) - 1).Caption = gchinas(k)
          End If
      End If
   If gfathers(k) = gchina(2) Then
       If gcols(k) <> 0 Then
           mnuadc(gcols(k) - 1).Caption = gchinas(k)
       End If
     End If
   If gfathers(k) = gchina(3) Then
       If gcols(k) <> 0 Then
           mnuadd(gcols(k) - 1).Caption = gchinas(k)
        End If
      End If
   If gfathers(k) = gchina(4) Then
       If gcols(k) <> 0 Then
         mnuade(gcols(k) - 1).Caption = gchinas(k)
       End If
      End If
   If gfathers(k) = gchina(5) Then
    If gcols(k) <> 0 Then
        mnuadf(gcols(k) - 1).Caption = gchinas(k)
    End If
     End If
   If gfathers(k) = gchina(6) Then
     If gcols(k) <> 0 Then
       mnuadg(gcols(k) - 1).Caption = gchinas(k)
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
   mnuad1.Caption = genglish(i)
Case 1:
    mnuad2.Visible = 1
   mnuad2.Caption = genglish(i)
Case 2:
   mnuad3.Visible = 1
   mnuad3.Caption = genglish(i)
Case 3:
   mnuad4.Visible = 1
   mnuad4.Caption = genglish(i)
Case 4:
   mnuad5.Visible = 1
   mnuad5.Caption = genglish(i)
Case 5:
   mnuad6.Visible = 1
   mnuad6.Caption = genglish(i)
Case 6
   mnuad7.Visible = 1
   mnuad7.Caption = genglish(i)
End Select
Next i
' 安置总子菜单
 For k = 0 To gsuncount - 1
   If gfathers(k) = gchina(0) Then
       mnuada(gcols(k) - 1).Caption = genglishs(k)
     End If
   If gfathers(k) = gchina(1) Then
       mnuadb(gcols(k) - 1).Caption = genglishs(k)
      End If
   If gfathers(k) = gchina(2) Then
       mnuadc(gcols(k) - 1).Caption = genglishs(k)
     End If
   If gfathers(k) = gchina(3) Then
       mnuadd(gcols(k) - 1).Caption = genglishs(k)
      End If
   If gfathers(k) = gchina(4) Then
       mnuade(gcols(k) - 1).Caption = genglishs(k)
      End If
   If gfathers(k) = gchina(5) Then
       mnuadf(gcols(k) - 1).Caption = genglishs(k)
     End If
   If gfathers(k) = gchina(6) Then
       mnuadg(gcols(k) - 1).Caption = genglishs(k)
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

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim FileNumber As Integer
Dim temp_china_english As Integer
     FileNumber = FreeFile
    Open "c:\ljxt\system.ini" For Input As #FileNumber
     Input #FileNumber, temp_china_english
     Input #FileNumber, Make_Mathine
    Close #FileNumber
  china_english = temp_china_english
Call first_do

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
            gnames(k) = recset.Fields("name").Value
            gfathers(k) = recset.Fields("father_name").Value
            gnexts(k) = recset.Fields("next_name").Value
            gprevs(k) = recset.Fields("prev_name").Value
            grows(k) = recset.Fields("row").Value
            gcols(k) = recset.Fields("col").Value
            gchinas(k) = recset.Fields("china_name").Value
            genglishs(k) = recset.Fields("english_name").Value
            gkeys(k) = recset.Fields("key").Value
            gflags(k) = recset.Fields("flag").Value
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








Private Sub Form_Unload(Cancel As Integer)
    comm_set_Form.sio_close_do
End Sub

Private Sub mnuada_Click(Index As Integer)
Select Case mnuada(Index).Caption
  Case Is = "退出":
                    Unload Me
                    End
  Case Is = "Exit":
                    Unload Me
                    End
  Case Is = "LIANJIAO":
                   MsgBox "This is a LIANJIAO contral oneself system", 1, "managment of database menu"
  Case Is = "生产命令录入":
                produce.Show
  Case Is = "生产命令运行":
                produce.Show
   Case Is = "生产命令模拟运行":
                produce.Show
End Select






 End Sub


Private Sub mnuadb_Click(Index As Integer)
Select Case mnuadb(Index).Caption
  Case Is = "材料录入":
                   Cai_liao_table_Input.Show
  Case Is = "碳黑录入":
                   Tan_hei_table_Input.Show
  Case Is = "油料录入":
                   You_liao_table_Input.Show
  Case Is = "胶料录入":
                   jiao_liao_table_Input.Show
  Case Is = "配方目录录入":
                   Pei_fan_table_Input.Show
  Case Is = "炼料录入":
                     lian_add_table_input.Show
                     lian_liao_table_input.Show
                   
                   
End Select

End Sub


Private Sub mnuadc_Click(Index As Integer)
Select Case mnuadc(Index).Caption
          Case Is = "电信号录入":
                        elec_disp.Show
          Case Is = "电信号运行":
                        elec_change.Show
                        
  End Select

End Sub


Private Sub mnuadd_Click(Index As Integer)
Select Case mnuadd(Index).Caption
          Case Is = "温度曲线图":
                        bl.Show
          Case Is = "power-tempro grap":
                        bl.Show
   End Select
End Sub


Private Sub mnuade_Click(Index As Integer)
Select Case mnuade(Index).Caption
  Case Is = "中英文切换":
         china_english = ENGLISH
         Call cmdenglish_Click
  Case Is = "菜单管理":
  
  Case Is = "背景设定":
  
  Case Is = "图示":
  
  Case Is = "通讯设置":
                     comm_set_Form.Show
  Case Is = "term_set":
                   comm_set_Form.Show
  Case Is = "设备参数":
                     comm_form.Show

  Case Is = "声卡测试":
            mci_Form.Show
  Case Is = "sound_test"
            mci_Form.Show
            
  Case Is = "数据传送"
             send_file_form.Show

  Case Is = "term_data":
                   comm_form.Show
                   
  Case Is = "CHINA-ENGLISH":
         china_english = CHINA
         Call cmdchina_Click
End Select
End Sub


Private Sub Timer1_Timer()
    Dim temp As Integer
    If comm_get_buffer.port < 1 Or comm_get_buffer.port > 8 Then
        Exit Sub
    End If
    temp = sio_getch(comm_get_buffer.port)
        If temp < 1 Then
            Exit Sub
        End If
    comm_get_buffer.head = comm_get_buffer.head + 1
    If comm_get_buffer.head > 1024 Then
        comm_get_buffer.head = 0
    End If
    comm_get_buffer.buffer(comm_get_buffer.head) = temp
End Sub



Private Sub first_do()
   Dim j As Integer
         For j = 1 To 20
              Comm_Port(j) = -1
         Next j
      
    
'     comm_set_Form.Form_Load
'     comm_set_Form.ok_command_Click
     If china_english = CHINA Then
             Me.Caption = "  炼胶系统    1997.7.13"
    Else
    
    End If
'    Timer1.Interval = 1
End Sub

Private Sub mnuadf_Click(Index As Integer)
Select Case mnuadf(Index).Caption
  Case Is = "故障表管理":
         error.Show
 
End Select

End Sub



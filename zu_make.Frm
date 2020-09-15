VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   1635
   ClientTop       =   1800
   ClientWidth     =   7365
   Height          =   4560
   Left            =   1575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7365
   Top             =   1455
   Width           =   7485
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\ljxt\make.mdb"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "make_name"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSRDC.MSRDC MSRDC1 
      Height          =   315
      Left            =   240
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   556
      _StockProps     =   4
      DataSourceName  =   ""
      RecordSource    =   ""
      LoginTimeout    =   15
      Connect         =   ""
      QueryTimeout    =   15
      RowsetSize      =   100
      RecordsetType   =   1
      MaxRows         =   -1
      LogMessages     =   ""
      KeysetSize      =   0
      ReadOnly        =   0   'False
      Options         =   0
      UserName        =   ""
      Password        =   ""
      CursorDriver    =   0
      EOFAction       =   0
      BOFAction       =   0
      ErrorThreshold  =   -1
      Caption         =   "MSRDC1"
      BackColor       =   16777215
      ForeColor       =   0
      Prompt          =   3
      LockType        =   3
      Appearance      =   1
   End
   Begin MSGrid.Grid Grd 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   5741
      _StockProps     =   77
      BackColor       =   16777215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim db1 As Database
Dim db2 As Database
Dim dy1 As Recordset
Dim dy2 As Recordset
Dim dy3 As Recordset
Dim td1 As TableDefs
Dim td2 As TableDefs
Dim fld As Fields
Dim fd As Fields
Dim zu As Integer
Dim zu10 As Integer
Dim zu11 As Integer
Dim zu20 As Integer
Dim zu21 As Integer
Dim zu30 As Integer
Dim zu31 As Integer
Dim zu40 As Integer
Dim zu41 As Integer
Dim zu50 As Integer
Dim zu51 As Integer
Dim zu60 As Integer
Dim zu61 As Integer

    
    
    





Private Sub Grid1_RowColChange()

End Sub

Private Sub Form_Load()
  
     Set db1 = OpenDatabase("c:\ljxt\make.mdb")
      Set db2 = OpenDatabase("c:\ljxt\make.mdb")
     Set dy1 = db1.OpenRecordset("make_name")
     Set dy2 = db2.OpenRecordset("make_lian_liao")
     Set dy3 = db2.OpenRecordset("make_base")
    
      zu10 = 0
      zu11 = 0
      zu20 = 0
      zu21 = 0
      zu30 = 0
      zu31 = 0
      zu40 = 0
      zu41 = 0
      zu50 = 0
      zu51 = 0
      zu60 = 0
      zu61 = 0
    Do Until dy2.EOF
        zu10 = zu10 + dy2.Fields("set_tan_time").Value
        zu11 = zu11 + dy2.Fields("get_tan_time").Value
        
        zu20 = zu10 + dy2.Fields("set_jiao_time").Value
        zu21 = zu11 + dy2.Fields("get_jiao_time").Value
        
        zu30 = zu10 + dy2.Fields("set_hun_time").Value
        zu31 = zu11 + dy2.Fields("get_hun_time").Value
        
        zu40 = zu10 + dy2.Fields("set_you_time").Value
        zu41 = zu11 + dy2.Fields("get_you_time").Value
        
        zu50 = zu10 + dy2.Fields("set_xiao_time").Value
        zu51 = zu11 + dy2.Fields("get_xiao_time").Value
        
        zu60 = zu10 + dy2.Fields("set_tan_time").Value
        zu61 = zu11 + dy2.Fields("get_tan_time").Value
         dy2.MoveNext
    Loop
       dy2.MoveFirst
     
       
     'dy2.MoveNext
   
     dy1.Index = "usidx"
     dy1.Seek "=", "投碳黑"
     If dy1.NoMatch Then
     MsgBox "did not find data1"
     Exit Sub
     End If
       dy1.Edit
        dy1.Fields("set_time") = zu10   'Value(db2.TableDefs(i).set_tan_time)
        dy1.Fields("get_time") = zu11      'dy2.Fields(10).Value
       dy1.Fields("get_tempro") = dy2.Fields(13).Value
       dy1.Fields("get_power") = dy2.Fields(14).Value
       dy1.Fields("set_tempro") = dy2.Fields(20).Value
       dy1.Fields("set_power") = dy2.Fields(21).Value
        dy1.Update
                        
           ' dy2.MoveNext
       
    
    
   
         
         dy1.Index = "usidx"
     dy1.Seek "=", "投胶料"
     If dy1.NoMatch Then
     MsgBox "did not find data2"
     Exit Sub
     End If
       dy1.Edit
        dy1.Fields("set_time") = zu20  'dy2.Fields(7).Value   'Value(db2.TableDefs(i).set_tan_time)
        dy1.Fields("get_time") = zu21 'dy2.Fields(8).Value
        dy1.Fields("get_tempro") = dy2.Fields(13).Value
        dy1.Fields("get_power") = dy2.Fields(14).Value
        dy1.Fields("set_tempro") = dy2.Fields(20).Value
       dy1.Fields("set_power") = dy2.Fields(21).Value
        dy1.Update
        
        
         dy1.Index = "usidx"
     dy1.Seek "=", "混料"
     If dy1.NoMatch Then
     MsgBox "did not find data3"
     Exit Sub
     End If
       dy1.Edit
        dy1.Fields("set_time") = zu30 ' dy2.Fields(3).Value   'Value(db2.TableDefs(i).set_tan_time)
        dy1.Fields("get_time") = zu31 ' dy2.Fields(4).Value
        dy1.Fields("get_tempro") = dy2.Fields(13).Value
        dy1.Fields("get_power") = dy2.Fields(14).Value
        dy1.Fields("set_tempro") = dy2.Fields(20).Value
       dy1.Fields("set_power") = dy2.Fields(21).Value
        dy1.Update
        
         dy1.Index = "usidx"
     dy1.Seek "=", "投油料"
     If dy1.NoMatch Then
     MsgBox "did not find data4"
     Exit Sub
     End If
       dy1.Edit
        dy1.Fields("set_time") = zu40 'dy2.Fields(5).Value   'Value(db2.TableDefs(i).set_tan_time)
        dy1.Fields("get_time") = zu41 'dy2.Fields(9).Value
        dy1.Fields("get_tempro") = dy2.Fields(13).Value
        dy1.Fields("get_power") = dy2.Fields(14).Value
       dy1.Fields("set_tempro") = dy2.Fields(20).Value
       dy1.Fields("set_power") = dy2.Fields(21).Value
       dy1.Update
        
        
         dy1.Index = "usidx"
     dy1.Seek "=", "投小料"
     If dy1.NoMatch Then
     MsgBox "did not find data5"
     Exit Sub
     End If
       dy1.Edit
        dy1.Fields("set_time") = zu50 'dy2.Fields(8).Value   'Value(db2.TableDefs(i).set_tan_time)
        dy1.Fields("get_time") = zu51 'dy2.Fields(12).Value
        dy1.Fields("get_tempro") = dy2.Fields(13).Value
        dy1.Fields("get_power") = dy2.Fields(14).Value
        dy1.Fields("set_tempro") = dy2.Fields(20).Value
       dy1.Fields("set_power") = dy2.Fields(21).Value
        dy1.Update
        
             dy2.MoveNext
              
        
        
      '  Do Until dy3.EOF
            dy1.Index = "usidx"
     dy1.Seek "=", "卸料"
     If dy1.NoMatch Then
     MsgBox "did not find data5"
     Exit Sub
     End If
       dy1.Edit
        dy1.Fields("set_time") = dy3.Fields("set_pai_time").Value   'Value(db2.TableDefs(i).set_tan_time)
        dy1.Fields("get_time") = dy3.Fields("get_pai_time").Value
         dy1.Fields("get_tempro") = dy3.Fields("pai_tempro").Value
        dy1.Fields("get_power") = dy3.Fields("pai_power")
        dy1.Fields("set_tempro") = dy2.Fields(20).Value
       dy1.Fields("set_power") = dy2.Fields(21).Value
        
        dy1.Update
       ' Debug.Print ; dy3.Fields("set_pai_time").Value
       ' Debug.Print ; dy3.Fields("get_pai_time").Value
     dy3.MoveNext
'Loop
   
     
     dy1.MoveFirst
       dy2.MoveFirst
       dy3.MoveFirst
      Set td2 = db2.TableDefs
    Set fd = td2("make_lian_liao").Fields
        i = 0
        zu = 0
   ' Do Until dy2.EOF
    ' Debug.Print ; dy3.Fields("set_pai_time").Value
    ' Debug.Print "tabledef:name;field:name "
   '  For i = 0 To fd.Count - 1
  ' Debug.Print "     ";
   ' Debug.Print "; "; dy2.Fields(13)
   ' Debug.Print "; "; dy2.Fields(14).Value
   'Debug.Print "; ";
    '  zu = zu + dy2.Fields("set_tan_time").Value
   ' Debug.Print "; "; zu + dy2.Fields("set_tan_time").Value
   'Next i
   
  ' dy2.MoveNext
   ' Loop
   ' Debug.Print "; "; zu
   ' Debug.Print "; "; zu + dy2.Fields("set_tan_time").Value
    ' Debug.Print "; "; zu + dy2.Fields("set_tan_time")
    ' Debug.Print dy2.Name
    '  Debug.Print dy2.RecordCount
    '  Debug.Print dy2.Updatable
     ' Debug.Print dy2.DateCreated
      'enumeratetabledef = True
    Set td1 = db1.TableDefs
    Set fld = td1("make_name").Fields
    
    Grd.FontSize = 12
    Grd.Cols = fld.Count
   ' MsgBox fld.Count
    For i = 0 To fld.Count - 1
        Grd.ColWidth(i) = Grd.Width / fld.Count
    Next i
    
    Grd.Row = 0
    For i = 0 To fld.Count - 1
       Grd.Col = i
       Grd.Text = fld(i).Name
    Next i
    
    i = 0
   ' dy1.MoveLast
   ' Grd.Rows = dy1.RecordCount + 1
  '  dy1.MoveFirst
    Do Until dy1.EOF
      i = i + 1
      Grd.Rows = i + 1
      Grd.Row = i
      For j = 0 To fld.Count - 1
              Grd.Col = j
              Grd.Text = dy1(j)
      Next j
       dy1.MoveNext
       Loop
     Grd.Refresh
       'dy1.Close
       'db1.Close
     

End Sub

Private Sub Grd1_RowColChange()

End Sub


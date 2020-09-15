VERSION 4.00
Begin VB.Form sound_from 
   Caption         =   "sound_form"
   ClientHeight    =   4155
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   Height          =   4560
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6690
   Top             =   1170
   Width           =   6810
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "write"
      Height          =   1095
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "read"
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "sound_from"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Const ChunkSize = 16384 ' Set size of chunk.
    Dim NumChunks As Integer, TotalSize As Long, X As Integer
    Dim MyDatabase As Database, MyRecordset As Recordset
    
    
    Set MyDatabase = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.MDB")
    Set MyRecordset = MyDatabase.OpenRecordset("error_num")
    
    If MyRecordset.RecordCount = 0 Then
        MyRecordset.Close
        MyDatabase.Close
'        Call speak_error("没记录")
        Exit Sub
    End If
    TotalSize = MyRecordset![china_speaker].FieldSize()
    If TotalSize = 0 Then
'        MyRecordset.Close
 '       MyDatabase.Close
        MsgBox "声音记录为空"
    End If
    
    TotalSize = FileLen("c:\ljxt\test.wav")
    NumChunks = TotalSize / ChunkSize - (TotalSize Mod ChunkSize <> 0)
    ReDim NoteArray(NumChunks) As String * ChunkSize
    MyRecordset.Edit
    Open "c:\ljxt\TEST.wav" For Binary As #1     ' Open file for output.
    For X = 1 To NumChunks
'        notearray(X) = MyRecordset![china_speaker].GetChunk((X - 1) * ChunkSize, ChunkSize)
        Get #1, , NoteArray(X)
        MyRecordset!china_speaker.AppendChunk NoteArray(X)
    Next X
    Close #1
     MyRecordset.Update
    MyRecordset.Close   ' Close table.
   MyDatabase.Close     'Close database.
End Sub


Private Sub Command2_Click()
Const ChunkSize = 16384 ' Set size of chunk.
    Dim NumChunks As Integer, TotalSize As Long, X As Integer
    Dim MyDatabase As Database, MyRecordset As Recordset
    Set MyDatabase = Workspaces(0).OpenDatabase("c:\ljxt\ljxt.MDB")
    ' Open table.
    Set MyRecordset = MyDatabase.OpenRecordset("error_num")
    
    ' Get field size.
    If MyRecordset.RecordCount = 0 Then
        MyRecordset.Close
        MyDatabase.Close
'        Call speak_error("没记录")
        Exit Sub
    End If
    TotalSize = MyRecordset![china_speaker].FieldSize()
        ' How many chunks?
    If TotalSize = 0 Then
        MyRecordset.Close
        MyDatabase.Close
'        Call speak_error("声音记录为空")
        Exit Sub
    End If
   NumChunks = TotalSize \ ChunkSize - (TotalSize Mod ChunkSize <> 0)
    Dim NoteArray    As String * ChunkSize
    ' Get current Notes field.
    
'    LpFormat.wBitsPerSample = 16
        
'    If waveOutOpen(LpWaveout, WAVE_MAPPER, LpFormat, 0, 0, CALLBACK_NULL) Then
 '           MsgBox "fail out waveform output "
  '          MyRecordset.Close
   '         MyDatabase.Close
    '        Exit Sub
   '     End If
    Open "c:\ljxt\TEST1.wav" For Binary As #1     ' Open file for output.
    For X = 1 To NumChunks
        NoteArray = MyRecordset![china_speaker].GetChunk((X - 1) * ChunkSize, ChunkSize)
        Put #1, , NoteArray
    Next X
    Close #1
    MyRecordset.Close   ' Close table.
   MyDatabase.Close     'Close database.
End Sub



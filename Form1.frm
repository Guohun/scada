VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2904
   ClientLeft      =   1068
   ClientTop       =   1428
   ClientWidth     =   5772
   Height          =   3288
   Left            =   1020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2904
   ScaleWidth      =   5772
   Top             =   1092
   Width           =   5868
   Begin VB.Timer get_time 
      Left            =   3720
      Top             =   1320
   End
   Begin VB.CommandButton send 
      Caption         =   "send"
      Height          =   852
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton get 
      Caption         =   "get"
      Height          =   372
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub get_Click()
                get_time.Interval = 5
End Sub

Private Sub get_time_Timer()
       Dim duan_hao As Integer
       Dim power  As Integer
       Dim tempro  As Integer
       Dim flag  As Integer
       Dim x As Integer

        x = read_power(duan_hao, power, tempro, flag)
         If flag = -1 Then
                Form1.Cls
         End If
         If flag = 1 Then
              Print power, tempro
         End If
End Sub

Private Sub send_Click()
Dim openforms
    For i = 1 To 1500 ' Start loop.
         Call write_power(1, i, i, 1)
         For j = 1 To 200
                 openforms = DoEvents    ' Yield to operating system.
         Next j
Next i  ' Increment loop counter.
End Sub


Private Sub Timer1_Timer()

End Sub



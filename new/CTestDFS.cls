VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestDFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'遍历图库
Public Function Normal_Traver()
Dim vi  As Integer
Dim Exit_Do As Boolean
Dim Visited_Vtx(1 To 100) As Integer

For vi = 1 To Pic_num
    Visited_Vtx(vi) = False:    Visited(vi) = False:    Passed_vtx(vi) = False
Next vi
For vi = 0 To Zline_num - 1
    Zline_obj_passing(vi) = False
Next vi
Do
    Exit_Do = True
    For vi = 1 To Pic_num
'        If (Not visited(vi)) And (dfs_is_start_vtx(vi) And (luo_firstadj(vi) <> -1)) Then
         If (Not Visited_Vtx(vi)) And (Not Passed_vtx(vi)) And (F_ShowLc.DFS_Is_Start_Vtx(vi) And (F_ShowLc.Luo_FirstAdj(vi) <> -1)) Then
            Visited_Vtx(vi) = True
            Last_Vtx = vi
            Road.Length = 0
            Call Normal_DFS(vi)
            Exit_Do = False
            Exit For
        End If
    Next vi
    Tmp = DoEvents
Loop Until Exit_Do

For vi = 0 To Zline_num - 1
    If Dev_ZLine(vi).Obj_Passing <> Zline_obj_passing(vi) Then
        Dev_ZLine(vi).Obj_Passing = Zline_obj_passing(vi)
    End If
Next vi
End Function

Public Function Normal_DFS(Cur_Vtx As Integer)
Dim i As Integer
Dim Next_Vtx As Integer
    'Visite (v0)
    Visited(Cur_Vtx) = True
    Next_Vtx = F_ShowLc.Luo_FirstAdj(Cur_Vtx)
    Last_Vtx = Cur_Vtx
    While Next_Vtx <> -1                   '当邻接点存在时
        If Not Visited(Next_Vtx) Then
            Call Normal_DFS(Next_Vtx)
        Else                'Rollback road
            If Arcs(Cur_Vtx, Next_Vtx) <> -1 Then
                Rollback_ZLine.ZIndex(Rollback_ZLine.Length) = Arcs(Cur_Vtx, Next_Vtx)
                Rollback_ZLine.Length = Rollback_ZLine.Length + 1
            End If
        End If
        Next_Vtx = F_ShowLc.Luo_NextAdj(Cur_Vtx, Next_Vtx) '找下一邻接点
        Last_Vtx = Cur_Vtx
    Wend
End Function

'遍历图库
Public Function Test_Traver()
Dim vi  As Integer
Dim Exit_Do As Boolean
Dim Visited_Vtx(1 To 100) As Integer

For vi = 1 To Pic_num
    Visited_Vtx(vi) = False:    Visited(vi) = False:    Passed_vtx(vi) = False
Next vi
For vi = 0 To Zline_num - 1
    Zline_obj_passing(vi) = False
Next vi
Do
    Exit_Do = True
    For vi = 1 To Pic_num
'        If (Not visited(vi)) And (dfs_is_start_vtx(vi) And (luo_firstadj(vi) <> -1)) Then
         If (Not Visited_Vtx(vi)) And (Not Passed_vtx(vi)) And (F_ShowLc.DFS_Is_Start_Vtx(vi) And (F_ShowLc.Luo_FirstAdj(vi) <> -1)) Then
            Visited_Vtx(vi) = True
            Last_Vtx = vi
            Road.Length = 0
            temp = 0
            Call Test_DFS(vi)
            Exit_Do = False
            Exit For
        End If
    Next vi
    Tmp = DoEvents
Loop Until Exit_Do

'环路遍历
Do
    Exit_Do = True
    For vi = 1 To Pic_num
         If (Not Visited_Vtx(vi)) And (Not Passed_vtx(vi)) And (F_ShowLc.Luo_FirstAdj(vi) <> -1) Then
            Passed_vtx(vi) = True
            Last_Vtx = vi
            Road.Length = 0
            temp = 0
            Call Test_DFS(vi)
            Exit_Do = False
            Exit For
        End If
    Next vi
    Tmp = DoEvents
Loop Until Exit_Do

For vi = 0 To Zline_num - 1
    If Dev_ZLine(vi).Obj_Passing <> Zline_obj_passing(vi) Then
        Dev_ZLine(vi).Obj_Passing = Zline_obj_passing(vi)
    End If
Next vi
End Function

Public Function Test_Set_Road()
Dim i As Integer
Static Road_Num As Integer
'If Road.length <= 0 Then
'    Exit Function
'End If

For i = 0 To Zline_num - 1
    Call F_ShowLc.Hide_Zline(i, True)
Next i
For i = 0 To Road.Length - 1
    Call F_ShowLc.Run_Zline(Road.ZIndex(i))
Next i
Road_Num = Road_Num + 1
MsgBox CStr(Road_Num)
End Function

'深度优先递归算法
Public Function Test_DFS(v0 As Integer)
Dim i As Integer
Dim W As Integer
    'Luo_Visite (v0)
    '记录折线路径
    If Arcs(Last_Vtx, v0) <> -1 Then
        Road.ZIndex(Road.Length) = Arcs(Last_Vtx, v0)
        Road.Length = Road.Length + 1
    End If
    Visited(v0) = True
    Passed_vtx(v0) = True
    '------------------------
    Last_Vtx = v0
    W = F_ShowLc.Luo_FirstAdj(v0)
    While W <> -1                   '当邻接点存在时
       If Not Visited(W) Then
            Search_new_road = True
            Last_Vtx = v0
            Call Test_DFS(W)
        Else        '存在环路
            If Arcs(v0, W) <> -1 Then
                Road.ZIndex(Road.Length) = Arcs(v0, W)
                Road.Length = Road.Length + 1
            End If
            MsgBox "Rollback Road"
            Call Test_Set_Road
        
            For k = 0 To Road.Length - 1
                If Dev_ZLine(Road.ZIndex(k)).End_Pic_Index = v0 Then
                    Road.Length = k + 1
                    Exit For
                End If
            Next k
            Visited(v0) = False
            Exit Function
        End If
        '找完一条路径,分析路径
        Last_Vtx = v0
        W = F_ShowLc.Luo_NextAdj(v0, W)  '找下一邻接点
    Wend
    '回搠
    If Search_new_road Then
        Call Test_Set_Road
        Search_new_road = False
    End If
    Visited(v0) = False
    If Road.Length > 0 Then
        Road.Length = Road.Length - 1
    End If
End Function

'Arc二维数组中的元素
'       -1                          无连接
'       0                           有连接(该元素不是产生环路的节点)
'       1 -- Pic_num        有连接(该元素是产生环路的节点)
Public Function Test_DFS_Zline_To_Arcs()
Dim i, j As Integer
For i = 1 To Pic_num
    For j = 1 To Pic_num
        Arcs(i, j) = -1
    Next j
Next i
For i = 0 To Zline_num - 1
    If Dev_ZLine(i).Style <> 11 Then   '不是供汽管
        Arcs(Dev_ZLine(i).Start_Pic_Index, Dev_ZLine(i).End_Pic_Index) = i
    End If
Next i
End Function


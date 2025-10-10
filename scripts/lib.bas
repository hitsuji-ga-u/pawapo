' libs >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' quick sort
' arrとして渡した配列を昇順にしたときのidxを受取る。
' [2 3 1] -> [3 1 2]を返す。
' 受ける配列はLong型にすること。
Function GetSortedIndexes(arr() As Double) As Long()
    Dim idx() As Long
    Dim n As Long, i As Long
    
    n = UBound(arr) - LBound(arr) + 1
    ReDim idx(LBound(arr) To UBound(arr))
    
    ' 元のインデックスを初期化
    For i = LBound(arr) To UBound(arr)
        idx(i) = i
    Next i
    
    ' ソート実行
    QuickSortCore arr, idx, LBound(arr), UBound(arr)
    
    GetSortedIndexes = idx
End Function

Private Sub QuickSortCore(arr() As Double, idx() As Long, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As Double
    Dim tmpD As Double, tmpI As Long
    
    i = first
    j = last
    pivot = (arr(first) + arr(last)) / 2   ' ピボットを両端の平均にする
    
    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            ' 値の交換
            tmpD = arr(i)
            arr(i) = arr(j)
            arr(j) = tmpD
            ' インデックスの交換
            tmpI = idx(i)
            idx(i) = idx(j)
            idx(j) = tmpI
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If first < j Then QuickSortCore arr, idx, first, j
    If i < last Then QuickSortCore arr, idx, i, last
End Sub

' find the index of the value in the array
Function FindIndex(arr() as Long, value) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            FindIndex = i
            Exit Function
        End If
    Next i

    FindIndex = -1
End Function


Function isArrayEmpty(arr_var As Variant)
    Dim p As Integer

    On Error Resume Next
        p = UBound(arr_var, 1)
    If Err.Number = 0 Then
        isArrayEmpty = False
    Else
        isArrayEmpty = True
    End If
End Function




Function GetShapeConers(shp As shape) As Variant
    ' 図形の4つの頂点座標を取得する
    ' ((left,       top),
    '  (left+width, top),
    '  (left+width, top+height),
    '  (left,       top+height))
    
    ' Args:
    '   shp (Shape): 頂点座標を取得する図形オブジェクト
    '
    ' Returns:
    '   Variant: 4頂点の座標を格納した2次元配列 (4x2)
    '           vertices(i,0): i番目の頂点のx座標
    '           vertices(i,1): i番目の頂点のy座標
    '
    ' Example:
    '   Dim vertices() As Long
    '   vertices = GetShapeConers(shp)
    '   For i = 0 To 3
    '       j = (i + 1) Mod 4
    '       shp1a(0) = vertices(i, 0) ' i番目の頂点のx座標
    '       shp1a(1) = vertices(i, 1) ' i番目の頂点のy座標
    '       shp1b(0) = vertices(j, 0) ' 次の頂点のx座標 
    '       shp1b(1) = vertices(j, 1) ' 次の頂点のy座標


    Dim vertices_0(3, 1) As Double
    Dim vertices(3, 1) As Double
    Dim center_x#, center_y#, s#, c#
    Dim i%

    center_x = CDbl(shp.left) + CDbl(shp.Width) / 2
    center_y = CDbl(shp.Top) + CDbl(shp.Height) / 2
    s = Sin(CDbl(shp.Rotation) * 3.14159265358979 / 180)
    c = Cos(CDbl(shp.Rotation) * 3.14159265358979 / 180)

    vertices_0(0, 0) = shp.left - center_x
    vertices_0(0, 1) = shp.Top - center_y
    vertices_0(1, 0) = shp.left + shp.Width - center_x
    vertices_0(1, 1) = shp.Top - center_y
    vertices_0(2, 0) = shp.left + shp.Width - center_x
    vertices_0(2, 1) = shp.Top + shp.Height - center_y
    vertices_0(3, 0) = shp.left - center_x
    vertices_0(3, 1) = shp.Top + shp.Height - center_y

    For i = 0 To 3
        vertices(i, 0) = vertices_0(i, 0) * c - vertices_0(i, 1) * s + center_x
        vertices(i, 1) = (vertices_0(i, 0) * s + vertices_0(i, 1) * c) + center_y
    Next

    GetShapeConers = vertices
End Function




Function is_crossed(Ax#, Ay#, Bx#, By#, Cx#, Cy#, Dx#, Dy#) As Boolean
    ' judgement that AB is clossing CD.
    ' return true when the other line is on the point B or D.

    Dim s#, t#

    s = (Cy - Ay) * (Bx - Ax) - (By - Ay) * (Cx - Ax)
    t = (Dy - Ay) * (Bx - Ax) - (By - Ay) * (Dx - Ax)

        If s * t > 0 Or s = 0 Then
        is_crossed = False
        Exit Function
    End If

    s = (Ay - Cy) * (Dx - Cx) - (Dy - Cy) * (Ax - Cx)
    t = (By - Cy) * (Dx - Cx) - (Dy - Cy) * (Bx - Cx)
        If s * t > 0 Or s = 0 Then
        is_crossed = False
        Exit Function
    End If

    is_crossed = True
End Function



Function nearest_node_index(shp As shape, x#, y#) As Long
    ' return the index of the nearest node from the argument point.
    Dim nearest_index&
    Dim shortest_distance#, distance#
    Dim i%

    nearest_index = 1
    shortest_distance = 999999
    For i = 1 To shp.Nodes.Count
        distance = (shp.Nodes(i).Points(1,1)-x) ^2 + (shp.Nodes(i).Points(1, 2) - y)^2
        if distance < shortest_distance then
            nearest_index = i
            shortest_distance = distance
        end if
    Next i
    nearest_node_index = nearest_index
End Function

' cast >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Function cm2pt(cm As double)
    cm2pt = cm / 0.0352777777777778
End Function

Function pt2cm(pt As Double)
    pt2cm = pt * 0.0352777777777778
End Function

' get a specific shape by name from shapes arg. >>>>>>>>>>>>>>>>
Function get_shape_by_name(shapes As Shapes, name As String) As Shape

    Dim shp As shape

    For Each shp In shapes
        If shp.Name = name Then
            set get_shape_by_name = shp
            Exit Function
        End If
    Next shp

    set get_shape_by_name = Nothing

End Function

Function min(ParamArray arglist()) As Double
    Dim i As Integer
    Dim min_val As Double
    
    min_val = arglist(LBound(arglist))
    For i = LBound(arglist) + 1 To UBound(arglist)
        If arglist(i) < min_val Then
            min_val = arglist(i)
        End If
    Next i

    min = min_val
End Function

Function max(ParamArray arglist()) As Double
    Dim i As Integer
    Dim max_val As Double
    
    max_val = arglist(LBound(arglist))
    For i = LBound(arglist) + 1 To UBound(arglist)
        If arglist(i) > max_val Then
            max_val = arglist(i)
        End If
    Next i
    max = max_val
End Function

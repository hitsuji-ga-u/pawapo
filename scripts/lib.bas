' libs >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' insertion sort 
Sub InsertionSortIndex(vals() As Double, indexes() As Integer)
    ' Doubleの配列varsの昇順で、indexesを並べ替える。
    Dim i&
    Dim j&
    Dim currentValue#
    Dim tmpIndex%

     For i = LBound(vals) + 1 To UBound(vals)
        currentValue = vals(i)
        j = i - 1
        tmpIndex = indexes(i)
        ' 適切な位置に要素を挿入する
        Do While j >= LBound(vals)
            If vals(j) > currentValue Then
                vals(j + 1) = vals(j)
                indexes(j + 1) = indexes(j)

            Else
                Exit Do
            End If
            j = j - 1
        Loop
        vals(j + 1) = currentValue
        indexes(j + 1) = tmpIndex
    Next i
End Sub


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
    ' example:
    ' Dim vertices() as Long
    ' vertices = GetShapeConers(shp)
    ' For i = 0 To 3
    '     j = (i + 1) Mod 4
    '     shp1a(0) = vertices(i, 0)
    '     shp1a(1) = vertices(i, 1)
    '     shp1b(0) = vertices(j, 0)
    '     shp1b(1) = vertices(j, 1)

    Dim vertices_0(3, 1) As Double
    Dim vertices(3, 1) As Double
    Dim Cx#, Cy#, s#, c#
    Dim i%

    Cx = CDbl(shp.left) + CDbl(shp.Width) / 2
    Cy = CDbl(shp.Top) + CDbl(shp.Height) / 2
    s = Sin(CDbl(shp.Rotation) * 3.14159265358979 / 180)
    c = Cos(CDbl(shp.Rotation) * 3.14159265358979 / 180)

    vertices_0(0, 0) = shp.left - Cx
    vertices_0(0, 1) = shp.Top - Cy
    vertices_0(1, 0) = shp.left + shp.Width - Cx
    vertices_0(1, 1) = shp.Top - Cy
    vertices_0(2, 0) = shp.left + shp.Width - Cx
    vertices_0(2, 1) = shp.Top + shp.Height - Cy
    vertices_0(3, 0) = shp.left - Cx
    vertices_0(3, 1) = shp.Top + shp.Height - Cy

    For i = 0 To 3
        vertices(i, 0) = vertices_0(i, 0) * c - vertices_0(i, 1) * s + Cx
        vertices(i, 1) = (vertices_0(i, 0) * s + vertices_0(i, 1) * c) + Cy
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

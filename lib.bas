
' 挿入ソート > > > > >> > > > > > > > > > > > > > > > >> > > >> >> > > > >
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



Function ShapeVertices(shp As Shape) As Variant
    '  時計回りで4角の頂点座標を 返却
    Dim vertices_0(3, 1) As Double
    Dim vertices(3, 1) As Double
    Dim cx#, cy#, s#, c#
    Dim i%
    
    
    cx = CDbl(shp.Left) + CDbl(shp.Width) / 2
    cy = shp.Top + shp.Height / 2
    s = Sin(shp.Rotation * 3.14159265358979 / 180)
    c = Cos(shp.Rotation * 3.14159265358979 / 180)
        
    vertices_0(0, 0) = shp.Left - cx
    vertices_0(0, 1) = shp.Top - cy
    vertices_0(1, 0) = shp.Left + shp.Width - cx
    vertices_0(1, 1) = shp.Top - cy
    vertices_0(2, 0) = shp.Left + shp.Width - cx
    vertices_0(2, 1) = shp.Top + shp.Height - cy
    vertices_0(3, 0) = shp.Left - cx
    vertices_0(3, 1) = shp.Top + shp.Height - cy

    For i = 0 To 3
        vertices(i, 0) = vertices_0(i, 0) * c - vertices_0(i, 1) * s + cx
        vertices(i, 1) = vertices_0(i, 0) * s + vertices_0(i, 1) * c + cy
    Next

    ShapeVertices = vertices
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


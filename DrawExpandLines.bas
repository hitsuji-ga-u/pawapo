
Dim DrawExpandLines()

    ' execute only when selecting 2 shapes
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange.Count = 2 Then Exit Sub

    If Not ActiveWindow.Selection.ShapeRange(1).Type = msoAutoShape And _
        Not ActiveWindow.Selection.ShapeRange(1).Type = msoPicture Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange(2).Type = msoAutoShape And _
        Not ActiveWindow.Selection.ShapeRange(2).Type = msoPicture Then Exit Sub

    Dim shp1 As Shape, shp2 As Shape
    Dim vertices() As Double
    Dim shp1a(1) As Double, shp1b(1) As Double, shp2a(1) As Double, shp2b(1) As Double
    Dim c1x#, c1y#, c2x#, c2y#

    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    Set shp2 = ActiveWindow.Selection.ShapeRange(2)

    ' when any shape type is micture, adding the shape on the picture which is the same size with the picture
    if shp1.type = msoPicture then
        Set shp1 = Activewindow.view.slide.Shapes.AddShape(msoShapeRectangle, shp1.left, shp1.Top, shp1.Width, shp1.Height)
    end if
    if shp2.type = msoPicture then
        Set shp2 = activewindow.slide.Shapes.AddShape(msoShapeRectangle, shp2.left, shp2.Top, shp2.Width, shp2.Height)
    end if

    c1x = CDbl(shp1.left) + CDbl(shp1.Width) / 2
    c1y = CDbl(shp1.Top) + CDbl(shp1.Height) / 2
    c2x = CDbl(shp2.left) + CDbl(shp2.Width) / 2
    c2y = CDbl(shp2.Top) + CDbl(shp2.Height) / 2

    vertices = ShapeVertices(shp1)

    Dim i&, j&
    Dim bl_is_crossed As Boolean
    bl_is_crossed = False
    
    For i = 0 To 3
        j = (i + 1) Mod 4
        shp1a(0) = vertices(i, 0)
        shp1a(1) = vertices(i, 1)
        shp1b(0) = vertices(j, 0)
        shp1b(1) = vertices(j, 1)

        If is_crossed(shp1a(0), shp1a(1), shp1b(0), shp1b(1), c1x, c1y, c2x, c2y) Then
            bl_is_crossed = True
            Exit For
        End If
    Next i

    vertices = ShapeVertices(shp2)
    bl_is_crossed = False
    For i = 0 To 3
        j = (i + 1) Mod 4
        shp2a(0) = vertices(i, 0)
        shp2a(1) = vertices(i, 1)
        shp2b(0) = vertices(j, 0)
        shp2b(1) = vertices(j, 1)

        If is_crossed(shp2a(0), shp2a(1), shp2b(0), shp2b(1), c1x, c1y, c2x, c2y) Then
            bl_is_crossed = True
            Exit For
        End If

    Next i

    ' set format
    shp1.Fill.Visible = msoFalse
    shp1.line.ForeColor.ObjectThemeColor = msoThemeColorAccent5
    shp1.line.Weight = 3
    shp2.Fill.Visible = msoFalse
    shp2.line.ForeColor.ObjectThemeColor = msoThemeColorAccent5
    shp2.line.Weight = 3

    ' drawing expansion line
    Dim ln1 As Shape
    Dim ln2 As Shape

    If is_crossed(shp1a(0), shp1a(1), shp2a(0), shp2a(1), shp1b(0), shp1b(1), shp2b(0), shp2b(1)) Then
        Set ln1 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1a(0), shp1a(1), shp2b(0), shp2b(1))
        Set ln2 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1b(0), shp1b(1), shp2a(0), shp2a(1))
    Else
        Set ln1 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1a(0), shp1a(1), shp2a(0), shp2a(1))
        Set ln2 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1b(0), shp1b(1), shp2b(0), shp2b(1))
    End If
    ln1.Line.Weight = 3
    ln1.Line.ForeCOlor.ObjectThemeColor = msoThemeColorAccent5
    ln1.Line.DashStyle = msoLineSysDot 
    ln2.Line.Weight = 3
    ln2.Line.ForeCOlor.ObjectThemeColor = msoThemeColorAccent5
    ln2.Line.DashStyle = msoLineSysDot 

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


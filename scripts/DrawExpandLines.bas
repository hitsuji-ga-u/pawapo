
SUb DrawExpandLines()

    ' execute only when selecting 2 shapes
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange.Count = 2 Then Exit Sub

    If Not ActiveWindow.Selection.ShapeRange(1).Type = msoAutoShape And _
        Not ActiveWindow.Selection.ShapeRange(1).Type = msoPicture And _
         Not Activewindow.selection.ShapeRange(1).Type = msoFreeform Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange(2).Type = msoAutoShape And _
        Not ActiveWindow.Selection.ShapeRange(2).Type = msoPicture And _
         Not Activewindow.selection.ShapeRange(2).Type = msoFreeform Then Exit Sub

    Dim shp1 As Shape, shp2 As Shape
    Dim vertices() As Double
    Dim shp1a(1) As Double, shp1b(1) As Double, shp2a(1) As Double, shp2b(1) As Double
    Dim c1x#, c1y#, c2x#, c2y#

    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    Set shp2 = ActiveWindow.Selection.ShapeRange(2)

    ' when any shape type is picture, adding the shape on the picture which is the same size with the picture
    if shp1.type = msoPicture then
        Set shp1 = Activewindow.view.slide.Shapes.AddShape(msoShapeRectangle, shp1.left, shp1.Top, shp1.Width, shp1.Height)
    end if
    if shp2.type = msoPicture then
        Set shp2 = activewindow.view.slide.Shapes.AddShape(msoShapeRectangle, shp2.left, shp2.Top, shp2.Width, shp2.Height)
    end if
    shp1.select
    shp2.select msoFalse

    ' calc each center of shapes
    c1x = CDbl(shp1.left) + CDbl(shp1.Width) / 2
    c1y = CDbl(shp1.Top) + CDbl(shp1.Height) / 2
    c2x = CDbl(shp2.left) + CDbl(shp2.Width) / 2
    c2y = CDbl(shp2.Top) + CDbl(shp2.Height) / 2


    Dim i&, j&

    vertices = GetShapeConers(shp1)
    For i = 0 To 3
        j = (i + 1) Mod 4
        shp1a(0) = vertices(i, 0)
        shp1a(1) = vertices(i, 1)
        shp1b(0) = vertices(j, 0)
        shp1b(1) = vertices(j, 1)

        If is_crossed(shp1a(0), shp1a(1), shp1b(0), shp1b(1), c1x, c1y, c2x, c2y) Then
            Exit For
        End If
    Next i

    vertices = GetShapeConers(shp2)
    For i = 0 To 3
        j = (i + 1) Mod 4
        shp2a(0) = vertices(i, 0)
        shp2a(1) = vertices(i, 1)
        shp2b(0) = vertices(j, 0)
        shp2b(1) = vertices(j, 1)

        If is_crossed(shp2a(0), shp2a(1), shp2b(0), shp2b(1), c1x, c1y, c2x, c2y) Then
            Exit For
        End If
    Next i

    ' set format
    shp1.Fill.Visible = msoFalse
    shp1.line.Weight = 2.25
    shp2.Fill.Visible = msoFalse
    shp2.line.Weight = 2.25

    ' add nodes
    AddNodes

    ' drawing expansion line
    Dim ln1 As Shape
    Dim ln2 As Shape
    dim connection_index&

    If is_crossed(shp1a(0), shp1a(1), shp2a(0), shp2a(1), shp1b(0), shp1b(1), shp2b(0), shp2b(1)) Then
        Set ln1 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1a(0), shp1a(1), shp2b(0), shp2b(1))
        connection_index = nearest_node_index(shp1, shp1a(0), shp1a(1))
        ln1.connectorformat.BeginConnect shp1, connection_index
        connection_index = nearest_node_index(shp2, shp2b(0), shp2b(1))
        ln1.connectorformat.EndConnect shp2, connection_index

        Set ln2 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1b(0), shp1b(1), shp2a(0), shp2a(1))
        connection_index = nearest_node_index(shp1, shp1b(0), shp1b(1))
        ln2.connectorformat.BeginConnect shp1, connection_index
        connection_index = nearest_node_index(shp2, shp2a(0), shp2a(1))
        ln2.connectorformat.EndConnect shp2, connection_index

    Else
        Set ln1 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1a(0), shp1a(1), shp2a(0), shp2a(1))
        connection_index = nearest_node_index(shp1, shp1a(0), shp1a(1))
        ln1.connectorformat.BeginConnect shp1, connection_index
        connection_index = nearest_node_index(shp2, shp2a(0), shp2a(1))
        ln1.connectorformat.EndConnect shp2, connection_index

        Set ln2 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1b(0), shp1b(1), shp2b(0), shp2b(1))
        connection_index = nearest_node_index(shp1, shp1b(0), shp1b(1))
        ln2.connectorformat.BeginConnect shp1, connection_index
        connection_index = nearest_node_index(shp2, shp2b(0), shp2b(1))
        ln2.connectorformat.EndConnect shp2, connection_index
    End If

    ln1.Line.Weight = 2.25
    ln1.Line.DashStyle = msoLineSysDot 
    ln2.Line.Weight = 2.25
    ln2.Line.DashStyle = msoLineSysDot 

    ln1.select msoFalse
    ln2.select msoFalse

End Sub




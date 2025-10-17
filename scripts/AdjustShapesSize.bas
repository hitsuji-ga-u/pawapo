

Sub AdjustShapesSize()
    ' adjust sizes of shapes 
    AdjustShapesWidth
    AdjustShapesHeight
End Sub

Sub AdjustShapesWidth()
    ' adjust width of shpaes to first selected shape

    ' only when selecting more than 1 shapes
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If
    If Not ActiveWindow.Selection.ShapeRange.Count >= 2 Then
        Exit Sub
    End If

    Dim shps As ShapeRange 
    Set shps = ActiveWindow.Selection.ShapeRange

    Dim base_shp As Shape
    Set base_shp = shps(shps.Count)


    ' get the width of the base shape
    Dim width_target#
    Dim coord() As Double
    coord = GetShapeConers(base_shp)
    width_target = max(coord(0, 0), coord(1, 0), coord(2, 0), coord(3, 0)) - min(coord(0, 0), coord(1, 0), coord(2, 0), coord(3, 0))

    Dim now_width#
    Dim delta#
    Dim dh#
    Dim dw#
    Dim rotation#

    Dim shp As Shape
    Dim i%
    For i = 1 To shps.Count - 1
        Set shp = shps(i)
        coord = GetShapeConers(shp)
        now_width = max(coord(0, 0), coord(1, 0), coord(2, 0), coord(3, 0)) - min(coord(0, 0), coord(1, 0), coord(2, 0), coord(3, 0))
        delta = width_target - now_width
        rotation = shp.Rotation
        dh = delta * Abs(Sin(rotation * 3.14159265358979 / 180))
        dw = delta * Abs(Cos(rotation * 3.14159265358979 / 180))
        shp.Width = shp.Width + dw
        shp.Height = shp.Height + dh
        debug.print shp.Name & ">  dw:" & dw & ", dh:" & dh & ", delta:" & delta 
    Next i
End Sub

Sub AdjustShapesHeight()
    ' adjust height of shapes to first selected shape

    ' only when selecting more than 1 shape
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If
    If Not ActiveWindow.Selection.ShapeRange.Count >= 2 Then
        Exit Sub
    End If

    Dim shps As ShapeRange 
    Set shps = ActiveWindow.Selection.ShapeRange

    Dim base_shp As Shape
    Set base_shp = shps(shps.Count)


    ' get the width of the base shape
    Dim height_target#
    Dim coord() As Double
    coord = GetShapeConers(base_shp)
    height_target = max(coord(0, 1), coord(1, 1), coord(2, 1), coord(3, 1)) - min(coord(0, 1), coord(1, 1), coord(2, 1), coord(3, 1))

    Dim now_height#
    Dim delta#
    Dim dh#
    Dim dw#
    Dim rotation#

    Dim shp As Shape
    Dim i%
    For i = 1 To shps.Count - 1
        Set shp = shps(i)
        coord = GetShapeConers(shp)
        now_height = max(coord(0, 1), coord(1, 1), coord(2, 1), coord(3, 1)) - min(coord(0, 1), coord(1, 1), coord(2, 1), coord(3, 1))
        delta = height_target - now_height
        rotation = shp.Rotation
        dh = delta * Abs(Cos(rotation * 3.14159265358979 / 180))
        dw = delta * Abs(Sin(rotation * 3.14159265358979 / 180))
        shp.Width = shp.Width + dw
        shp.Height = shp.Height + dh
        debug.print shp.Name & ">  dw:" & dw & ", dh:" & dh & ", delta:" & delta 
    Next i
End Sub


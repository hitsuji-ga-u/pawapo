

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

    Dim shp1 As Shape
    Dim shp As Shape
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = shp1.Width
    Next shp
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

    Dim shp1 As Shape
    Dim shp As Shape
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Height = shp1.Height
    Next shp
End Sub


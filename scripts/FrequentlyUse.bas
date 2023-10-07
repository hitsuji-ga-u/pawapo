
Sub FrequentlyArrowStyle()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As Shape

    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoLine Or shp.AutoShapeType = msoShapeMixed Then
            shp.Line.EndArrowheadLength = msoArrowheadLong
            shp.Line.EndArrowheadWidth = msoArrowheadWide
            shp.Line.EndArrowheadStyle = msoArrowheadOpen
            shp.Line.Weight = 3
        End If
    Next

End Sub


Sub FrequentlyShadeStyle()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As Shape

    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoLine Or shp.AutoShapeType = msoShapeMixed Then
            shp.Line.EndArrowheadLength = msoArrowheadLong
            shp.Line.EndArrowheadWidth = msoArrowheadWide
            shp.Line.EndArrowheadStyle = msoArrowheadOpen
            shp.Line.Weight = 3
        End If
    Next

End Sub

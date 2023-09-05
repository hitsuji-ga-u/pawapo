
Sub FrequentlyUse()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As Shape

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    Debug.Print shp.Type
    Debug.Print shp.AutoShapeType


    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoLine Or shp.AutoShapeType = msoShapeMixed Then
            
            shp.Line.EndArrowheadLength = msoArrowheadLong
            shp.Line.EndArrowheadWidth = msoArrowheadWide
            shp.Line.EndArrowheadStyle = msoArrowheadOpen
            shp.Line.Weight = 1
        End If
    Next

End Sub



Sub FrequentlyArrowStyle15()
    FrequentlyArrowStyle(1.5)
End Sub
Sub FrequentlyArrowStyle30()
    FrequentlyArrowStyle(3)
End Sub


Sub FrequentlyArrowStyle(width As Double)
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As Shape

    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoLine Or shp.AutoShapeType = msoShapeMixed Then
            shp.Line.EndArrowheadLength = msoArrowheadLong
            shp.Line.EndArrowheadWidth = msoArrowheadWide
            shp.Line.EndArrowheadStyle = msoArrowheadOpen
            shp.Line.Weight = width
        End If
    Next

End Sub

Sub FrequentlyShadowStyleOn()
    ' when only selecting shps
    if not activewindow.selection.type = ppSelectionShapes then exit sub

    Dim shp As shape

    for each shp in ActiveWindow.selection.ShapeRange
        shp.Shadow.Visible = False
    next shp

End Sub

Sub FrequentlyShadowStyleOff()
    ' when only selecting shps
    if not activewindow.selection.type = ppSelectionShapes then exit sub

    Dim shp As shape

    for each shp in ActiveWindow.selection.ShapeRange
        shp.Shadow.Visible = True
        shp.Shadow.Style = msoShadowStyleOuterShadow
        shp.Shadow.Blur = 5 ' ぼかし半径
        shp.Shadow.Transparency = 0.6
        shp.Shadow.OffsetX = 2.121319152764454 ' X方向のオフセット
        shp.Shadow.OffsetY = 2.121319152764454 ' Y方向のオフセット
    next shp

End Sub

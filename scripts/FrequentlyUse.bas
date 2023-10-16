

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
        If shp.Type = msoLine Or shp.Type = msoFreeform Or shp.AutoShapeType = msoShapeMixed Then
            shp.line.EndArrowheadLength = msoArrowheadLong
            shp.line.EndArrowheadWidth = msoArrowheadWide
            shp.line.EndArrowheadStyle = msoArrowheadOpen
            shp.line.Weight = width
        End If
    Next
End Sub


Sub FrequentlyShadowStyleOff()
    ' when only selecting shps
    if not activewindow.selection.type = ppSelectionShapes then exit sub

    Dim shp As shape

    for each shp in ActiveWindow.selection.ShapeRange
        shp.Shadow.Visible = False
    next shp

End Sub


Sub FrequentlyShadowStyleOn()
    ' when only selecting shps
    if not activewindow.selection.type = ppSelectionShapes then exit sub

    Dim shp As shape

    for each shp in ActiveWindow.selection.ShapeRange
        shp.Shadow.Visible = True
        shp.Shadow.Style = msoShadowStyleOuterShadow
        shp.Shadow.Blur = 5 ' Blur radius
        shp.Shadow.Transparency = 0.6
        shp.Shadow.OffsetX = 2.121319152764454 ' x-offset
        shp.Shadow.OffsetY = 2.121319152764454 ' y-offset
    next shp

End Sub

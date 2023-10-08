
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
    ' set shadow format
    Dim slide As slide
    Set slide = ActivePresentation.slides(1) ' スライド番号を指定

    ' when only selecting shps
    if not activewindow.selection.type = ppSelectionShapes then exit sub

    Dim shp As shape

    for each shp in ActiveWindow.selection.ShapeRange
        shp.Shadow.Visible = True
        shp.Shadow.Style = msoShadowStyleOuterShadow
        shp.Shadow.Blur = 5 ' ぼかし半径
        shp.Shadow.Transparency = 0.6
        shp.Shadow.OffsetX = 10 ' X方向のオフセット
        shp.Shadow.OffsetY = 10 ' Y方向のオフセット
    next shp

End Sub
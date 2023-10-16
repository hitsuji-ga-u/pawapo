Sub test()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As shape

    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoLine Or shp.Type = msoFreeform Or shp.AutoShapeType = msoShapeMixed Then
            shp.line.EndArrowheadLength = msoArrowheadLong
            shp.line.EndArrowheadWidth = msoArrowheadWide
            shp.line.EndArrowheadStyle = msoArrowheadOpen
            shp.line.Weight = 1.5
        End If
    Next
End Sub


sub test1()
    Dim shp1 As shape
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    debug.print shp1.rotation

    shp1.Shadow.Visible = True
    shp1.Shadow.Style = msoShadowStyleOuterShadow
    shp1.Shadow.Blur = 5 ' ぼかし半径
    shp1.Shadow.Transparency = 0.6
    shp1.Shadow.OffsetX = 10 ' X方向のオフセット
    shp1.Shadow.OffsetY = 10 ' Y方向のオフセット
    shp1.Shadow.Obscured = msoFalse
        
end sub




Sub FrequentlyArrowStyle15()
    FrequentlyArrowStyle(1.5)
End Sub
Sub FrequentlyArrowStyle30()
    FrequentlyArrowStyle(3)
End Sub

Sub FrequentlyArrowStyleBoth()
    On Error Resume Next
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As Shape

    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoLine Or shp.Type = msoFreeform Or shp.AutoShapeType = msoShapeMixed Then
            shp.line.EndArrowheadLength = msoArrowheadLong
            shp.line.EndArrowheadWidth = msoArrowheadWide
            shp.line.EndArrowheadStyle = msoArrowheadOpen
            shp.line.BeginArrowheadLength = msoArrowheadLong
            shp.line.BeginArrowheadWidth = msoArrowheadWide
            shp.line.BeginArrowheadStyle = msoArrowheadOpen
        End If

    Next
End Sub


Sub FrequentlyArrowStyle(width As Double)
    On Error Resume Next
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As Shape

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.line.Weight = width
        If shp.Type = msoLine Or shp.Type = msoFreeform Or shp.AutoShapeType = msoShapeMixed Then
            shp.line.EndArrowheadLength = msoArrowheadLong
            shp.line.EndArrowheadWidth = msoArrowheadWide
            shp.line.EndArrowheadStyle = msoArrowheadOpen
        End If
    Next
End Sub

Sub FrequentlyLineStyleON()
    ' when only selecting shps
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As shape

    FrequentlyShadowStyleOn

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Fill.Visible = msoFalse
        shp.line.Weight = 3#
    Next shp

End Sub

Sub FrequentlyShadowStyleOff()
    ' when only selecting shps
    if not activewindow.selection.type = ppSelectionShapes then exit sub

    Dim shp As shape

    for each shp in ActiveWindow.selection.ShapeRange
        shp.Shadow.Visible = msoFalse
    next shp

End Sub


Sub FrequentlyShadowStyleOn()
    ' when only selecting shps
    if not activewindow.selection.type = ppSelectionShapes then exit sub

    Dim shp As shape

    for each shp in ActiveWindow.selection.ShapeRange
        if shp.Type = msoTextBox Then
            shp.TextFrame.TextRange.Font.Shadow = msoTrue
            shp.TextFrame.TextRange2.Font.Shadow.Type = msoShadow21 
            shp.TextFrame.TextRange.Font.Shadow.Blur = 3 ' Blur radius
            shp.TextFrame.TextRange.Font.Shadow.Transparency = 0.6
            shp.TextFrame.TextRange.Font.Shadow.OffsetX = 2.121319152764454 ' x-offset
            shp.TextFrame.TextRange.Font.Shadow.OffsetY = 2.121319152764454 ' y-offset
        else
            shp.Shadow.Visible = msoTrue
            shp.Shadow.Style = msoShadowStyleOuterShadow
            shp.Shadow.Blur = 4 ' Blur radius
            shp.Shadow.Transparency = 0.6
            shp.Shadow.OffsetX = 2.121319152764454 ' x-offset
            shp.Shadow.OffsetY = 2.121319152764454 ' y-offset
        End if
    next shp

End Sub



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
            shp.TextFrame2.TextRange.Font.Shadow.Visible = msoTrue
            shp.TextFrame2.TextRange.Font.Shadow.Type = msoShadow21 
            shp.TextFrame2.TextRange.Font.Shadow.Blur = 3 ' Blur radius
            shp.TextFrame2.TextRange.Font.Shadow.Transparency = 0.6
            shp.TextFrame2.TextRange.Font.Shadow.OffsetX = 2.121319152764454 ' x-offset
            shp.TextFrame2.TextRange.Font.Shadow.OffsetY = 2.121319152764454 ' y-offset
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


Sub CrossSymbol()

    Dim symbol_size#, start_x#, start_y#, end_x#, end_y#
    symbol_size = cm2pt(1)
    start_x = ActiveWindow.Presentation.PageSetup.SlideWidth / 2 - symbol_size / 2
    start_y = ActiveWindow.Presentation.PageSetup.SlideHeight / 2 - symbol_size / 2
    end_x = ActiveWindow.Presentation.PageSetup.SlideWidth / 2 + symbol_size / 2
    end_y = ActiveWindow.Presentation.PageSetup.SlideHeight / 2 + symbol_size / 2

    Dim line1 As shape, line2 As shape

    Set line1 = ActiveWindow.Selection.SlideRange(1).Shapes.AddLine(start_x, start_y, end_x, end_y)
    Set line2 = ActiveWindow.Selection.SlideRange(1).Shapes.AddLine(start_x, end_y, end_x, start_y)

    With line1.line
        .Weight = 3
        .ForeColor.RGB = RGB(255, 0, 0)
    End With
    With line2.line
        .Weight = 3
        .ForeColor.RGB = RGB(255, 0, 0)
    End With

    line1.Select
    line2.Select msoFalse

    ActiveWindow.Selection.ShapeRange.Group.select

End Sub


Sub CircleSymbol()

    Dim symbol_size#, start_x#, start_y#, end_x#, end_y#
    symbol_size = cm2pt(1)
    start_x = ActiveWindow.Presentation.PageSetup.SlideWidth / 2 - symbol_size / 2
    start_y = ActiveWindow.Presentation.PageSetup.SlideHeight / 2 - symbol_size / 2
    end_x = ActiveWindow.Presentation.PageSetup.SlideWidth / 2 + symbol_size / 2
    end_y = ActiveWindow.Presentation.PageSetup.SlideHeight / 2 + symbol_size / 2

    Dim c As shape

    With ActiveWindow.Selection.SlideRange(1).Shapes.AddShape(msoShapeOval, start_x, start_y, symbol_size, symbol_size)
        .Select
        .Fill.Visible = msoFalse
        .line.ForeColor.RGB = RGB(0, 0, 255)
        .line.Weight = 3
    End With

End Sub


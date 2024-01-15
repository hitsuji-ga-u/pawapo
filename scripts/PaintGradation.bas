
' change a shape fill to gradation color  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub PaintGradation()

    On Error Resume Next
    Debug.Print ActiveWindow.Selection.Type; ppSelectionText

    if not ActiveWindow.selection.type = ppSelectionShapes then exit sub

    dim tgt_shp as Shape

    For Each tgt_shp In ActiveWindow.Selection.ShapeRange
        tgt_shp.line.Visible = msoFalse
        tgt_shp.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
        tgt_shp.Fill.OneColorGradient msoGradientHorizontal, 1, 1
        tgt_shp.Fill.GradientStops(1).Color.ObjectThemeColor = msoThemeLight1
        tgt_shp.Fill.GradientStops(1).Transparency = 1
        tgt_shp.Fill.GradientStops(2).Position = 0.9
    Next tgt_shp
End Sub

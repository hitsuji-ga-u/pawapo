
' change a shape fill to gradation color  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub PaintGradation()

    On Error Resume Next
    Debug.Print ActiveWindow.Selection.Type; ppSelectionText

    if not ActiveWindow.selection.type = ppSelectionShapes then exit sub

    dim tgt_shp as Shape

    For each tgt_shp in Activewindow.selection.ShapeRange
        tgt_shp.Line.Visible = msoFalse
        tgt_shp.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
        tgt_shp.Fill.OneColorGradient msoGradientHorizontal, 2, 1
    next tgt_shp

End Sub

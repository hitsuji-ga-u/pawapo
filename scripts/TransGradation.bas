' 透明グラデーションをつける  > > > > > > > > > > > > >> > > > > > > > > > > > > > > >
Sub TransGradation()
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub
    With ActiveWindow.Selection.ShapeRange(1)
        .Line.Visible = msoFalse
        With .Fill
            .ForeColor.ObjectThemeColor = msoThemeLight1
            .OneColorGradient msoGradientHorizontal, 3, 1
            .GradientStops(1).Transparency = 1
            .GradientStops(2).Position = 0.6
            .GradientStops(2).Transparency = 0.3
        End With
    End With
End Sub


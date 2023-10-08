Sub test()

    Dim shp As shape

    for each shp in ActiveWindow.selection.ShapeRange
        shp.Shadow.Visible = True
        shp.Shadow.Style = msoShadowStyleOuterShadow
        shp.Shadow.Blur = 5 ' ぼかし半径
        shp.Shadow.Transparency = 0.6
        shp.Shadow.OffsetX = 10 ' X方向のオフセット
        shp.Shadow.OffsetY = 10 ' Y方向のオフセット
        shp.Shadow.Obscured = msoFalse
    next shp

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


Sub test()

End Sub

sub test1()
    Dim shp1 As shape
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    debug.print shp1.rotation

end sub


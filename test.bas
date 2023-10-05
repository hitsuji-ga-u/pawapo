Sub test()

End Sub

sub test1()



    Dim shp1 As shape
    Dim shp2 As shape
    
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    Set shp2 = ActiveWindow.Selection.ShapeRange(2)
    shp1.connectformat.BeginConnect shp2, 1
end sub


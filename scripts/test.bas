' test >>>>> test >>>>> test >>>>> test >>>>> test >>>>> test >>>>> test >>>>> test >>>>>
Sub test()

    ' execute only when selecting 2 shapes
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange.Count = 2 Then Exit Sub

    If Not ActiveWindow.Selection.ShapeRange(1).Type = msoAutoShape And _
        Not ActiveWindow.Selection.ShapeRange(1).Type = msoPicture And _
         Not Activewindow.selection.ShapeRange(1).Type = msoFreeform Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange(2).Type = msoAutoShape And _
        Not ActiveWindow.Selection.ShapeRange(2).Type = msoPicture And _
         Not Activewindow.selection.ShapeRange(2).Type = msoFreeform Then Exit Sub

    Dim shp1 As Shape, shp2 As Shape
    Dim vertices() As Double
    Dim shp1a(1) As Double, shp1b(1) As Double, shp2a(1) As Double, shp2b(1) As Double
    Dim c1x#, c1y#, c2x#, c2y#

    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    Set shp2 = ActiveWindow.Selection.ShapeRange(2)

    ' when any shape type is picture, adding the shape on the picture which is the same size with the picture
    if shp1.type = msoPicture then
        Set shp1 = Activewindow.view.slide.Shapes.AddShape(msoShapeRectangle, shp1.left, shp1.Top, shp1.Width, shp1.Height)
    end if
    if shp2.type = msoPicture then
        Set shp2 = activewindow.view.slide.Shapes.AddShape(msoShapeRectangle, shp2.left, shp2.Top, shp2.Width, shp2.Height)
    end if
    shp1.select
    shp2.select msoFalse


End Sub


Sub test1()
    Dim lay As CustomLayout
    Dim shp As shape

    For Each lay In ActivePresentation.SlideMaster.CustomLayouts
        If lay.name = "タイトルのみ" Then
            Debug.Print "lay: "; lay.name

            For Each shp In lay.shapes
                Debug.Print shp.name
          Next shp

        End If
    Next lay


    Dim title As shape
    For Each lay In ActivePresentation.SlideMaster.CustomLayouts
        Set title = get_shape_by_name(lay.shapes, "Placeholder 1")
        If title Is Nothing Then Exit Sub
    Next lay
End Sub
    
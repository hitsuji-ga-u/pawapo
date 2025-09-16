' Align Center >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub AlignCenterVertical()
    ' vertically align the centers of selected shapes with the last shape.

    ' no selecting
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim shps As ShapeRange

    set shps = ActiveWindow.Selection.ShapeRange

    ' if selected only 1 shape, align the center of the shape to the center of the slide
    If shps.Count = 1 Then
        shps.Align msoAlignMiddles, msoTrue

    ' if selected more than 1 shape, align the centers of the shapes to the center of the last shape
    Elseif shps.Count >= 2 Then
        Dim i&

        for i = 1 To shps.Count - 1
            shps(i).Top = shps(shps.Count).Top + shps(shps.Count).Height/2 - shps(i).Height / 2
        next i
    end If
End sub

Sub AlignCenterHorizontal()
    ' horizontally align the centers of selected shapes with the last shape.

    ' no selecting
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim shps As ShapeRange

    set shps = ActiveWindow.Selection.ShapeRange

    ' if selected only 1 shape, align the center of the shape to the center of the slide
    If shps.Count = 1 Then
        shps.Align msoAlignCenters, msoTrue

    ' if selected more than 1 shape, align the centers of the shapes to the center of the last shape
    Elseif shps.Count >= 2 Then
        Dim i&

        for i = 1 To shps.Count - 1
            shps(i).Left = shps(shps.Count).Left + shps(shps.Count).Width/2 - shps(i).Width / 2
        next i
    end If
End sub

Sub AlignCenter()
    AlignCenterHorizontal
    AlignCenterVertical
End sub


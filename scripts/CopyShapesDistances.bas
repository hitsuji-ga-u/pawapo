

' 図形間の距離をコピー、ペースト > > > > >> > > > >> > > >> > > >> > > >> > > >> > >> 
Sub CopyShapeDistances()
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then Exit Sub
    
    Dim shp1 As Shape, shp2 As Shape
    
    With ActiveWindow.Selection
        If .ShapeRange(1).Top < .ShapeRange(2).Top Then
            Set shp1 = .ShapeRange(1)
            Set shp2 = .ShapeRange(2)
        Else
            Set shp1 = .ShapeRange(2)
            Set shp2 = .ShapeRange(1)
        End If
    End With

    ShapeDistanceY = shp2.Top - shp1.Top - shp1.Height


    With ActiveWindow.Selection
        If .ShapeRange(1).left < .ShapeRange(2).left Then
            Set shp1 = .ShapeRange(1)
            Set shp2 = .ShapeRange(2)
        Else
            Set shp1 = .ShapeRange(2)
            Set shp2 = .ShapeRange(1)
        End If
    End With
    
    ShapeDistanceX = shp2.left - shp1.left - shp1.Width
End Sub



' 図形間の距離ペースト Y方向 > > > > >> > > > >> > 

Sub PasteShpaeDistancesX()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then Exit Sub
    
    Dim i&
    
    For i = 2 To ActiveWindow.Selection.ShapeRange.Count
        With ActiveWindow.Selection
            .ShapeRange(i).Left = .ShapeRange(i - 1).Left + .ShapeRange(i - 1).Width + ShapeDistanceX
        End With
    Next i

End Sub

Sub PasteShpaeDistancesY()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then Exit Sub
    
    Dim i&
    
    For i = 2 To ActiveWindow.Selection.ShapeRange.Count
        With ActiveWindow.Selection
            .ShapeRange(i).Top = .ShapeRange(i - 1).Top + .ShapeRange(i - 1).Height + ShapeDistanceY
        End With
    Next i

End Sub
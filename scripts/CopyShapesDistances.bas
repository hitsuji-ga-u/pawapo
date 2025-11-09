




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
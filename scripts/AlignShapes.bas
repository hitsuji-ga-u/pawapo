
Sub ObjectsAlignTopLeft()
    ' align shapes
    If not activewindow.selection.type = ppSelectionShapes then exit sub
    CommandBars.ExecuteMso "ObjectsAlignLeftSmart"
    CommandBars.ExecuteMso "ObjectsAlignTopSmart"
End Sub

Sub ObjectsAlignTopRight()
    ' align shapes
    If not activewindow.selection.type = ppSelectionShapes then exit sub
    CommandBars.ExecuteMso "ObjectsAlignRightSmart"
    CommandBars.ExecuteMso "ObjectsAlignTopSmart"
End Sub

Sub ObjectsAlignBottomLeft()
    ' align shapes
    If not activewindow.selection.type = ppSelectionShapes then exit sub
    CommandBars.ExecuteMso "ObjectsAlignLeftSmart"
    CommandBars.ExecuteMso "ObjectsAlignBottomSmart"
End Sub

Sub ObjectsAlignBottomRight()
    ' align shapes
    If not activewindow.selection.type = ppSelectionShapes then exit sub
    CommandBars.ExecuteMso "ObjectsAlignRightSmart"
    CommandBars.ExecuteMso "ObjectsAlignBottomSmart"
End Sub

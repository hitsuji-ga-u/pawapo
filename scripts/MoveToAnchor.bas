

' move selected shape to anchor position >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub MoveToAnchor()

    Debug.Print ActiveWindow.Selection.Type; ppSelectionText

    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        Exit Sub
    End If

    ' get slide size
    dim w#, h#
    w = ActivePresentation.PageSetup.SlideWidth
    h = ActivePresentation.PageSetup.SlideHeight

    ' get anchor pos
    Dim x#, y#
    ' x = w * 0.02
    x = h * 0.02
    y = h * 0.02 + cm2pt(1.59)

    ' set shape to anchor
    ActiveWindow.Selection.ShapeRange(1).left = x
    ActiveWindow.Selection.ShapeRange(1).Top = y

End Sub

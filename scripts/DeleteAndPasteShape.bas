

' delete selected shape and paste  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub DeleteAndPasteShape()
    On Error Resume Next

    ' delete selected shapes
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ActiveWindow.Selection.ShapeRange.Delete
    End If

    ' paste from clipboard
    ActiveWindow.View.Paste

End Sub
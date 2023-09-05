

' 図形削除 & ペースト >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub DeleteAndPasteShape()
    On Error Resume Next

    ' 図形選択していたら削除
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ActiveWindow.Selection.ShapeRange.Delete
    End If

    ' コピーしている図形をペースト
    ActiveWindow.View.Paste

    Exit Sub
ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume Next
End Sub
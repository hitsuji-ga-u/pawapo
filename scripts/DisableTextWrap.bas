Sub DisableTextWrap()
    ' 図形内で改行しないにチェックしたテキストボックスを挿入する
    ' あるいは、選択した図形の図形内で改行をしないにチェックをいれる

    On Error GoTo ErrorHandler

    ' 何も選択してない場合
    If ActiveWindow.Selection.Type = ppSelectionNone Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Dim textbox As Shape

        Set textbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        ActiveWindow.Presentation.PageSetup.SlideWidth / 2, _
                        ActiveWindow.Presentation.PageSetup.SlideHeight / 2, 0, 0)
        textbox.Select
        textbox.TextFrame.TextRange.Text = ""
    End If

    If ActiveWindow.Selection.Type = ppSelectionText Then
        If ActiveWindow.Selection.TextRange.Parent.Parent.HasTextFrame Then
            ActiveWindow.Selection.TextRange.Parent.Parent.TextFrame2.WordWrap = msoFalse
        End If

    ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim selectedTextBox As Shape

            Set selectedTextBox = ActiveWindow.Selection.ShapeRange(1)

        For Each selectedTextBox In ActiveWindow.Selection.ShapeRange
            If selectedTextBox.HasTextFrame Then
                selectedTextBox.TextFrame2.WordWrap = msoFalse
            End If

            Next selectedTextBox

    End If

    Exit Sub
ErrorHandler:
    Resume Next
End Sub
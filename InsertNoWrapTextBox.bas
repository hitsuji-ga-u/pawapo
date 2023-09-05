
' テキストボックス挿入 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub InsertNoWrapTextBox()
    ' 図形内で改行しないにチェックしたテキストボックスを挿入する
    ' あるいは、選択した図形の図形内で改行をしないにチェックをいれる

    ' 何も選択してない場合挿入する。余白0。折り返ししないチェックは以降の処理で行う。
    If ActiveWindow.Selection.Type = ppSelectionNone Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Dim textbox As Shape

        Set textbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        ActiveWindow.Presentation.PageSetup.SlideWidth / 2, _
                        ActiveWindow.Presentation.PageSetup.SlideHeight / 2, 0, 0)

        textbox.TextFrame.DeleteText
        textbox.TextFrame.TextRange.Select
    End If

    ' テキスト選択中の場合、折り返ししないにチェックする
    If ActiveWindow.Selection.Type = ppSelectionText Then
        If ActiveWindow.Selection.TextRange.Parent.Parent.HasTextFrame Then
            With ActiveWindow.Selection.TextRange.Parent.Parent.TextFrame
                .WordWrap = msoFalse
                .MarginTop = 0
                .MarginRight = 0
                .MarginBottom = 0
                .MarginLeft = 0
            End With
        End If

    ' 1つ以上の図形を選択中の場合、すべての図形で折り返ししないにチェックする
    ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim selectedTextBox As Shape

        For Each selectedTextBox In ActiveWindow.Selection.ShapeRange
            If selectedTextBox.HasTextFrame Then
                With selectedTextBox.TextFrame
                    .WordWrap = msoFalse
                    .MarginTop = 0
                    .MarginRight = 0
                    .MarginBottom = 0
                    .MarginLeft = 0
                End With
            End If
        Next selectedTextBox
    End If

    Exit Sub
ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume Next
End Sub

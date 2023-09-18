
' insert textbox >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub InsertNoWrapTextBox()
    ' insert no wrap textbox or make textbox to no wrap textbox

    ' no selecting
    If ActiveWindow.Selection.Type = ppSelectionNone Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Dim textbox As Shape

        Set textbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        ActiveWindow.Presentation.PageSetup.SlideWidth / 2, _
                        ActiveWindow.Presentation.PageSetup.SlideHeight / 2, 0, 0)

        textbox.TextFrame.DeleteText
        textbox.TextFrame.TextRange.Select
        textbox.TextRange.Font.Size = 16
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

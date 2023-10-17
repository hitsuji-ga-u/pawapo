
' insert textbox >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub InsertNoWrapTextBox()
    ' insert no wrap textbox or make textbox to no wrap textbox

    ' when selecting nothing, selecting a table, selecting a img,  Insert a txt box.
    If activewindow.selection.type = ppSelectionShapes then
        dim shp as shape
        set shp = activewindow.selection.ShapeRange(1)
            if shp.type = msoTable Or shp.type = msoPicture then
                AddTextbox
            End If
    End If
    If ActiveWindow.Selection.Type = ppSelectionNone Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        AddTextbox
    End If

    ' when selecting txt
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

    ' when selecting shapes
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
End Sub

Sub AddTextbox()
    Dim textbox As Shape

    Set textbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                    msoTextOrientationHorizontal, _
                    ActiveWindow.Presentation.PageSetup.SlideWidth / 2, _
                    ActiveWindow.Presentation.PageSetup.SlideHeight / 2, 0, 0)

    textbox.TextFrame.DeleteText
    textbox.TextFrame.TextRange.Select
    textbox.textframe.TextRange.Font.Size = 16
End Sub

' 表の幅を文字に合わせる       >>>> > > > > >> > > > > > > >> > > > >> > > >> > > >> >
Sub TableWidthAutoFit()

    If ActiveWindow.Selection.Type = ppSelectionNone Then Exit Sub
    If ActiveWindow.Selection.Type = ppSelectionSlides Then Exit Sub
    If not ActiveWindow.Selection.ShapeRange(1).Type = msoTable Then Exit Sub
    
    ' テキストボックスを使って文字サイズをはかる。
    ' テキスト、フォント、文字サイズ、を合わせる。
    Dim i_col&, i_row&

    Dim table As table
    Debug.Print ActiveWindow.Selection.ShapeRange(1).Type

    Set table = ActiveWindow.Selection.ShapeRange(1).table

    Dim txtbox As Shape
    Dim horizontal_margin&
    Dim max_width&

    For i_col = 1 To table.Columns.Count
        max_width = 0
        For i_row = 1 To table.Rows.Count
            With table.Cell(i_row, i_col).Shape.TextFrame
                horizontal_margin = .MarginLeft + .MarginRight

            End With
            Set txtbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                            msoTextOrientationHorizontal, 0, 0, 0, 0)
            txtbox.TextFrame.WordWrap = msoFalse
            txtbox.TextFrame.Orientation = table.Cell(i_row, i_col).Shape.TextFrame.Orientation
            With txtbox.TextFrame.TextRange
                .Text = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Text
                With .Font
                    .Name = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Font.Name
                    .Bold = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Font.Bold
                    .Italic = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Font.Italic
                    .Size = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Font.Size
                End With

                If max_width < .BoundWidth + horizontal_margin Then max_width = .BoundWidth + horizontal_margin
            End With
            txtbox.Delete
            table.Columns(i_col).Width = max_width
        Next i_row
    Next i_col
End Sub

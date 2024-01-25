' tuggle title alignment >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub TuggleTitleAlignment()
    Dim s As slide
    Dim title As shape
    Dim determined_flg As Boolean
    Dim dst_alignment As PpParagraphAlignment

    determined_flg = False

    For Each s In ActivePresentation.Slides
        Set title = get_shape_by_name(s.shapes, "Title 1")
        Debug.Print "a"
        If title Is Nothing Then GoTo continue

        Debug.Print dst_alignment
        If Not determined_flg Then
            If Not title.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft Then
                dst_alignment = ppAlignLeft
            Else
                dst_alignment = ppAlignCenter
            End If
            determined_flg = True
        End If
        Debug.Print dst_alignment
        title.TextFrame.TextRange.ParagraphFormat.Alignment = dst_alignment
continue:
    Next s
End Sub

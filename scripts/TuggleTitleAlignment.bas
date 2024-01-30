' tuggle title alignment >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub TuggleTitleAlignment()
    Dim s As slide
    Dim title As shape
    Dim determined_flg As Boolean
    Dim dst_alignment As PpParagraphAlignment

    determined_flg = False

    For Each s In ActivePresentation.Slides
        Set title = get_shape_by_name(s.shapes, "Title 1")

        If title Is Nothing Then GoTo continue

        If Not determined_flg Then
            If Not title.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft Then
                dst_alignment = ppAlignLeft
            Else
                dst_alignment = ppAlignCenter
            End If
            determined_flg = True
        End If
        title.TextFrame.TextRange.ParagraphFormat.Alignment = dst_alignment

continue:
    Next s


    Dim lay As CustomLayout
    For Each lay In ActivePresentation.SlideMaster.CustomLayouts
        If lay.name = "タイトルのみ" Then
            Set title = get_shape_by_name(lay.shapes, "タイトル プレースホルダー 1")
            If Not title Is Nothing Then
                title.TextFrame.TextRange.ParagraphFormat.Alignment = dst_alignment
            End If 
        End If
    next lay

End Sub

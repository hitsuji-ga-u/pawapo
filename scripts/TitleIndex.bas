Sub TitleIndex()
    Dim sld As slide
    Dim title As shape
    Dim regex As Object
    Dim start_sld As Long
    Dim title_txt_pre As String
    Dim title_txt As String
    Dim bl_multiple As Boolean
    Dim i&

    bl_multiple = False
    title_txt_pre = ""

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\(\d*/\d*\)"

    For Each sld In ActivePresentation.Slides
        Set title = get_shape_by_name(sld.shapes, "Title 1")

        If Not title Is Nothing Then
            ' 　タイトル空白が連続しても同じとみなさない。別処理。
            If title.TextFrame.HasText Then
                title_txt = Trim(regex.Replace(title.TextFrame.TextRange.text, ""))
                title.TextFrame.TextRange.text = title_txt

                ' 同じタイトルの開始
                If Not bl_multiple Then
                    If title_txt = title_txt_pre Then
                        start_sld = sld.SlideIndex - 1
                        bl_multiple = True
                    End If
                Else
                ' 同じタイトルの終了
                    If title_txt <> title_txt_pre Then
                        For i = start_sld To sld.SlideIndex - 1
                            get_shape_by_name(ActivePresentation.Slides(i).shapes, "Title 1").TextFrame.TextRange.InsertAfter (" (" & CStr(i - start_sld + 1) & "/" & CStr(sld.SlideIndex - start_sld) & ")")
                        Next
                        bl_multiple = False
                    End If
                End If

               If sld.SlideIndex = ActivePresentation.Slides.Count Then
                    If bl_multiple Then
                        For i = start_sld To ActivePresentation.Slides.Count
                            get_shape_by_name(ActivePresentation.Slides(i).shapes, "Title 1").TextFrame.TextRange.InsertAfter (" (" & CStr(i - start_sld + 1) & "/" & CStr(sld.SlideIndex - start_sld + 1) & ")")
                        Next
                        bl_multiple = False
                    End If
                End If
            Else
            ' タイトルが空白はタイトル終了と同値
                If bl_multiple Then
                    For i = start_sld To sld.SlideIndex - 1
                        get_shape_by_name(ActivePresentation.Slides(i).shapes, "Title 1").TextFrame.TextRange.InsertAfter (" (" & CStr(i - start_sld + 1) & "/" & CStr(sld.SlideIndex - start_sld) & ")")
                    Next
                    bl_multiple = False
                End If
            End If
            title_txt_pre = title_txt
        Else
            If bl_multiple Then
                For i = start_sld To sld.SlideIndex - 1
                    get_shape_by_name(ActivePresentation.Slides(i).shapes, "Title 1").TextFrame.TextRange.InsertAfter (" (" & CStr(i - start_sld + 1) & "/" & CStr(sld.SlideIndex - start_sld) & ")")
                Next
                bl_multiple = False
            End If
            title_txt_pre = ""
        End If
    Next sld
End Sub
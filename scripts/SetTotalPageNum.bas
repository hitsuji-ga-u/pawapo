' setting total page. need to add a shape named "total_page" to Slidemaster
' total_page and edit_text are loaded at initialization

Sub SetTotalSlidNumber(page As Long)
    Dim shp As Shape

    set shp = get_shape_by_name(ActivePresentation.SlideMaster.shapes, "page_index")

    If shp Is Nothing Then
        msgbox "Please set the name of the text box representing the page number to ""page_index""."
        Exit Sub
    End If

    shp.TextFrame.TextRange.text = ""
    shp.TextFrame.TextRange.InsertSlideNumber
    shp.TextFrame.TextRange.InsertAfter ("/" & CStr(page))

End Sub

Sub SetPageEditBox(control As IRibbonControl, ByRef text)
    edit_text = text
    text = CStr(total_page)
End Sub

Sub SetTotalPageNum(control As IRibbonControl)

    total_page = ActivePresentation.Slides.Count - 1
    SetTotalSlidNumber total_page
    edit_text = CStr(total_page)
    ribbon.InvalidateControl("total_page")
End Sub

Sub RefleshTotalPageNum(control As IRibbonControl, ByRef text)
    ' if input not numerical value, undo.
    if not isnumeric(text) Then
        text = CStr(total_page)
        ribbon.InvalidateControl("total_page")
        Exit Sub
    End If

    total_page = CLng(text)
    text = CStr(total_page)

    SetTotalSlidNumber total_page

End Sub

' getting total page. if page_num has already set, return set num.
' It is needed that the textbox which shows the page-num is set its name as "page_index".
Function GetNowTotalPage() As Long
    Dim page_num&
    Dim page_num_txtbox$
    page_num_txtbox = "page_index"
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = ".#./[\d]{1,}"

    Dim shp As shape
    Set shp = get_shape_by_name(ActivePresentation.SlideMaster.shapes, page_num_txtbox)

    If shp Is Nothing Then GetNowTotalPage = ActivePresentation.Slides.Count - 1: Exit Function

    If Not regex.test(shp.TextFrame.TextRange.text) Then GetNowTotalPage = ActivePresentation.Slides.Count - 1: Exit Function

    Dim matches As Object
    Set matches = regex.Execute(shp.TextFrame.TextRange.text)

    regex.Pattern = "\d+(?=$)"
    Set matches = regex.Execute(matches(0).Value)
    page_num = matches(0).Value

    GetNowTotalPage = page_num
End Function
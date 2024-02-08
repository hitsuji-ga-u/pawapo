' margin setting of textbox >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'margin_horizontal, margin_vertical are loaded at initialization

Sub GetMarginHorizontal(control As IRibbonControl, ByRef text)
   text = CStr(pt2cm(margin_horizontal))
end Sub

Sub GetMarginVertical(control As IRibbonControl, ByRef text)
    text = CStr(pt2cm(margin_vertical))
End Sub

Sub SetMarginHorizontal(control As IRibbonControl, ByRef text As String)
    if not isnumeric(text) Then
        text = CStr(margin_horizontal)
        ribbon.InvalidateControl("margin_horizontal")
        Exit Sub
    End If

    margin_horizontal = cm2pt(CDbl(text))
End Sub

Sub SetMarginVertical(control As IRibbonControl, ByRef text As String)
    if not isnumeric(text) Then
        text = CStr(margin_vertical)
        ribbon.InvalidateControl("margin_vertical")
        Exit Sub
    End If

    margin_vertical = cm2pt(CDbl(text))
End Sub

Sub ApplyMargin()
    ' テキストボックス、図形、表で適用できるように。
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    ' 表にもShpaeがある。そのTextFrameからmarginをいじれる
    Dim shp As Shape
    Dim txtfrm As TextFrame

    ' 表の場合
    ' 表でない場合
    For each shp In ActiveWindow.selection.ShapeRange
        If shp.Type = msoTable Then
        Else 
            shp.TextFrame.MarginLeft = margin_horizontal
            shp.TextFrame.MarginRight = margin_horizontal
            shp.TextFrame.MarginTop = margin_vertical
            shp.TextFrame.MarginBottom = margin_vertical
        End If 
    Next shp

End Sub
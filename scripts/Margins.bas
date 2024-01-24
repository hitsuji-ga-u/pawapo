' margin setting of textbox >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'margin_horizontal, margin_vertical are loaded at initialization

Sub GetMarginHorizontal(control As IRibbonControl, ByRef text)
   text = CStr(margin_horizontal)
end Sub

Sub GetMarginVertical(control As IRibbonControl, ByRef text)
    text = CStr(margin_vertical)
End Sub

Sub SetMarginHorizontal(control As IRibbonControl, ByRef text As String)
    MsgBox text
End Sub

Sub SetMarginVertical(control As IRibbonControl, ByRef text As String)
    MsgBox text
End Sub

Sub ApplyMargin()
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

End Sub
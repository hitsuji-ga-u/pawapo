

' 図形を白のグラデーションにする　 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub PaintGradation()

    Debug.Print ActiveWindow.Selection.Type; ppSelectionText

    if not ActiveWindow.selection.type = ppSelectionShapes then exit sub

    dim tgt_shp as Shape

    ' 線を無しにする
    tgt_shp.Line.Visible = msoFalse

    ' テーマカラーの1色目を塗りつぶしに使用する
    shape.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
    shape.Fill.ForeColor.Brightness = 0

End Sub

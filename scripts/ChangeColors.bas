
' 図形の塗りつぶし色、枠線の色、フォントの色を変える > > > >> > > > >> > > > >> > > >

Sub ChangeTextColorLight1()
    ChangeTextColor msoThemeColorLight1
end sub
Sub ChangeTextColorDark1()
    ChangeTextColor msoThemeColorDark1
end sub
Sub ChangeTextColorDark2()
    ChangeTextColor msoThemeColorDark2
end sub
Sub ChangeTextColorAccent2()
    ChangeTextColor msoThemeColorAccent2
end sub
Sub ChangeTextColorAccent3()
    ChangeTextColor msoThemeColorAccent3
end sub
Sub ChangeTextColorAccent4()
    ChangeTextColor msoThemeColorAccent4
end sub
Sub ChangeTextColorAccent5()
    ChangeTextColor msoThemeColorAccent5
end sub
Sub ChangeTextColorAccent6()
    ChangeTextColor msoThemeColorAccent6
end sub
Sub ChangeTextColorRed()
    ChangeTextColor 0, 255, 0, 0
end sub

Sub ChangeLineColorLight1()
    ChangeLineColor msoThemeColorLight1
end sub
Sub ChangeLineColorDark1()
    ChangeLineColor msoThemeColorDark1
end sub
Sub ChangeLineColorDark2()
    ChangeLineColor msoThemeColorDark2
end sub
Sub ChangeLineColorAccent2()
    ChangeLineColor msoThemeColorAccent2
end sub
Sub ChangeLineColorAccent3()
    ChangeLineColor msoThemeColorAccent3
end sub
Sub ChangeLineColorAccent4()
    ChangeLineColor msoThemeColorAccent4
end sub
Sub ChangeLineColorAccent5()
    ChangeLineColor msoThemeColorAccent5
end sub
Sub ChangeLineColorAccent6()
    ChangeLineColor msoThemeColorAccent6
end sub
Sub ChangeLineColorRed()
    ChangeLineColor 0, 255, 0, 0
end sub
Sub ChangeLineColorNone()
    ChangeLineColor -1
end sub

Sub ChangeShapeColorLight1()
    ChangeShapeColor msoThemeColorLight1
end sub
Sub ChangeShapeColorDark1()
    ChangeShapeColor msoThemeColorDark1
end sub
Sub ChangeShapeColorDark2()
    ChangeShapeColor msoThemeColorDark2
end sub
Sub ChangeShapeColorAccent2()
    ChangeShapeColor msoThemeColorAccent2
end sub
Sub ChangeShapeColorAccent3()
    ChangeShapeColor msoThemeColorAccent3
end sub
Sub ChangeShapeColorAccent4()
    ChangeShapeColor msoThemeColorAccent4
end sub
Sub ChangeShapeColorAccent5()
    ChangeShapeColor msoThemeColorAccent5
end sub
Sub ChangeShapeColorAccent6()
    ChangeShapeColor msoThemeColorAccent6
end sub
Sub ChangeShapeColorRed()
    ChangeShapeColor 0, 255, 0, 0
end sub
Sub ChangeShapeColorNone()
    ChangeShapeColor -1 
end sub



Sub ChangeShapeColor(color_idx As Long, Optional r As Long = 0, Optional g As Long = 0, Optional b As Long = 0)
    ' change fill color of shapes
    ' color_idx: 
    '     specify msoThemeColor
    '     specify 0 to specify RGB
    '     specify -1 for no fill
    If ActiveWindow.Selection.Type = ppSelectionShapes Then 
        Dim i&
        Dim shp As Shape
        For Each shp In ActiveWindow.Selection.ShapeRange
            If color_idx = 0 Then
                shp.Fill.Visible = msoTrue
                shp.Fill.ForeColor.RGB = RGB(r, g, b)
            Elseif color_idx = -1 Then
                shp.Fill.Visible = msoFalse
            Else
                shp.Fill.Visible = msoTrue
                shp.Fill.ForeColor.ObjectThemeColor = color_idx
            End If
        Next shp

    ElseIf ActiveWindow.selection.type = ppSelectionText then
        Set shp = ActiveWindow.selection.Textrange.parent.parent

        If color_idx = 0 Then
            shp.Fill.Visible = msoTrue
            shp.Fill.ForeColor.RGB = RGB(r, g, b)
        Elseif color_idx = -1 Then
            shp.Fill.Visible = msoFalse
        Else
            shp.Fill.Visible = msoTrue
            shp.Fill.ForeColor.ObjectThemeColor = color_idx
        End If
    End If
End Sub

Sub ChangeTextColor(color_idx As Long, Optional r As Long = 0, Optional g As Long = 0, Optional b As Long = 0)
    If ActiveWindow.Selection.Type = ppSelectionShapes Then

        Dim i&
        Dim shp As Shape
        For Each shp In ActiveWindow.Selection.ShapeRange
            If color_idx = 0 Then
                shp.TextFrame.TextRange.Font.Color.RGB = RGB(r, g, b)
            Else
                shp.TextFrame.TextRange.Font.Color.ObjectThemeColor = color_idx
            End If
        Next shp
    ElseIf ActiveWindow.selection.type = ppselectiontext then
        Dim txtrange As TextRange
        set txtrange = ActiveWindow.selection.Textrange

        If color_idx = 0 Then
            TxtRange.Font.Color.RGB = RGB(r, g, b)
        Else
            txtrange.Font.Color.ObjectThemeColor = color_idx
        End If
    End If
End Sub

Sub ChangeLineColor(color_idx As Long, Optional r As Long = 0, Optional g As Long = 0, Optional b As Long = 0)
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim i&
        Dim shp As Shape
        For Each shp In ActiveWindow.Selection.ShapeRange
            if color_idx = 0 Then
                shp.Line.Visible = msoTrue
                shp.Line.ForeColor.RGB = RGB(r, g, b)
            Elseif color_idx = -1 Then
                shp.Line.Visible = msoFalse
            else
                shp.Line.Visible = msoTrue
                shp.Line.ForeColor.ObjectThemeColor = color_idx
            end if
        Next shp
    ElseIf ActiveWindow.selection.type = ppselectiontext then
        Set shp = ActiveWindow.selection.Textrange.parent.parent

        if color_idx = 0 Then
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = RGB(r, g, b)
        Elseif color_idx = -1 Then
            shp.Line.Visible = msoFalse
        else
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.ObjectThemeColor = color_idx
        end if
    End if
End sub

' Align Center >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub AlignCenterVertical()
    ' 1つめに選択した図形の中央に合わせる　上下中央

    ' 図形を選択してなければ終わり
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim shps As ShapeRange

    set shps = ActiveWindow.Selection.ShapeRange

    ' 1のみ選択の場合
    If shps.Count = 1 Then
        shps.Align msoAlignMiddles, msoTrue

    ' 2つ以上選択している場合
    Elseif shps.Count >= 2 Then
        Dim i&

        for i = 2 To shps.Count
            shps(i).Top = shps(1).Top + shps(1).Height/2 - shps(i).Height / 2
        next i
    end If
End sub

Sub AlignCenterHorizontal()
    ' 1つめに選択した図形の中央に合わせる　左右中央

    ' 図形を選択してなければ終わり
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim shps As ShapeRange

    set shps = ActiveWindow.Selection.ShapeRange

    ' 1のみ選択の場合
    If shps.Count = 1 Then
        shps.Align msoAlignCenters, msoTrue

    ' 2つ以上選択している場合
    Elseif shps.Count >= 2 Then
        Dim i&

        for i = 2 To shps.Count
            shps(i).Left = shps(1).Left + shps(1).Width/2 - shps(i).Width / 2
        next i
    end If
End sub

Sub AlignCenter()
    AlignCenterHorizontal
    AlignCenterVertical
End sub


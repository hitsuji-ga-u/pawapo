' 図形の位置コピー >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub CopyShapesPos()
    ' 図形の位置を格納する

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim selectedShapes As ShapeRange
    Dim i As Long

    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    ReDim shapePositions(1 To selectedShapes.Count, 1 To 2) ' 2次元配列 (x, y)

    For i = 1 To selectedShapes.Count
        shapePositions(i, 1) = selectedShapes(i).left
        shapePositions(i, 2) = selectedShapes(i).Top
    Next i
End Sub

Sub PasteShapesAbsolutely()
    ' 選択した図形をコピーしてある位置に絶対的に合わせる

    ' 位置コピーされてなければ終了
    If isArrayEmpty(shapePositions) Then
        Exit Sub
    End If

    ' 図形選択されてなければ終了
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim i&
    Dim selectedShpsNum As Integer

    ' 選択された図形の数を取得
    selectedShpsNum = ActiveWindow.Selection.ShapeRange.Count

    ' min(図形の選択数, コピーしてある図形の位置数)個の図形を調整する。
    For i = 1 To IIf(UBound(shapePositions) < selectedShpsNum, UBound(shapePositions), selectedShpsNum)
        With ActiveWindow.Selection.ShapeRange(i)
            .left = shapePositions(i, 1)
            .Top = shapePositions(i, 2)
        End With
    Next i
End Sub

Sub PasteShapesRelatively()
    ' 選択した図形をコピーしてある位置に相対的に合わせる

    ' 位置コピーされてなければ終了
    If isArrayEmpty(shapePositions) Then
        Exit Sub
    End If

    ' 図形選択されてなければ終了
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    ' 位置コピー数が2以上なければ終了
    If UBound(shapePositions) - LBound(shapePositions) + 1 < 2 Then
        Exit Sub
    End If

    ' 図形二つ以上選択されていなければ終了
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        Exit Sub
    End If
        Dim i&

    Dim selectedShpsNum As Integer

    ' 選択された図形の数を取得
    selectedShpsNum = ActiveWindow.Selection.ShapeRange.Count

    ' min(図形の選択数, コピーしてある図形の位置数)個の図形を調整する。
    For i = 2 To IIf(UBound(shapePositions) < selectedShpsNum, UBound(shapePositions), selectedShpsNum)
        With ActiveWindow.Selection
            .ShapeRange(i).left = .ShapeRange(1).left + shapePositions(i, 1) - shapePositions(1, 1)
            .ShapeRange(i).Top = .ShapeRange(1).Top + shapePositions(i, 2) - shapePositions(1, 2)
        End With
    Next i
End Sub

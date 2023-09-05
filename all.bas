
Option Explicit

Dim shapePositions() As Variant



Sub AdjustShapesHeight()
    ' 図形の高さ揃える
        
    ' Shape選択中判定
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If
    If Not ActiveWindow.Selection.ShapeRange.Count >= 2 Then
        Exit Sub
    End If
    
     
    Dim shp1 As Shape
    Dim shp As Shape
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Height = shp1.Height
    Next shp
    
End Sub


Sub AdjustShapesSize()
    ' 図形の大きさ揃える
        
    AdjustShapesWidth
    AdjustShapesHeight
    
End Sub



Sub AdjustShapesWidth()
    ' 図形の幅揃える
        
    ' Shape選択中判定
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If
    If Not ActiveWindow.Selection.ShapeRange.Count >= 2 Then
        Exit Sub
    End If
    
     
    Dim shp1 As Shape
    Dim shp As Shape
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)

    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.Width = shp1.Width
    Next shp
    
End Sub





' Align Center >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub AlignCenterVertical()
    ' 1つめに選択した図形の中央に合わせる　上下中央

    ' 図形を選択してなければ終わり
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim shps As ShapeRange

    Set shps = ActiveWindow.Selection.ShapeRange

    ' 1のみ選択の場合
    If shps.Count = 1 Then
        shps.Align msoAlignMiddles, msoTrue

    ' 2つ以上選択している場合
    ElseIf shps.Count >= 2 Then
        Dim i&

        For i = 2 To shps.Count
            shps(i).Top = shps(1).Top + shps(1).Height / 2 - shps(i).Height / 2
        Next i
    End If
End Sub

Sub AlignCenterHorizontal()
    ' 1つめに選択した図形の中央に合わせる　左右中央

    ' 図形を選択してなければ終わり
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim shps As ShapeRange

    Set shps = ActiveWindow.Selection.ShapeRange

    ' 1のみ選択の場合
    If shps.Count = 1 Then
        shps.Align msoAlignCenters, msoTrue

    ' 2つ以上選択している場合
    ElseIf shps.Count >= 2 Then
        Dim i&

        For i = 2 To shps.Count
            shps(i).left = shps(1).left + shps(1).Width / 2 - shps(i).Width / 2
        Next i
    End If
End Sub

Sub AlignCenter()
    AlignCenterHorizontal
    AlignCenterVertical
End Sub














' 図形をくっつけて並べる >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub AlignShapesHorizontalStick()
    ' 図形をくっつけて並べる　横

    ' 2個以上のShape選択中判定
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim numShapes%

            numShapes = ActiveWindow.Selection.ShapeRange.Count

        If numShapes >= 2 Then
            Dim shp1, shp2 As Shape
            Dim i%
            Dim lefts() As Double
            Dim indexes() As Integer
            ReDim lefts(1 To numShapes)
            ReDim indexes(1 To numShapes)

            For i = 1 To numShapes
                lefts(i) = ActiveWindow.Selection.ShapeRange(i).left
                indexes(i) = i
            Next i

            InsertionSortIndex lefts, indexes

            For i = 1 To numShapes - 1

                Set shp1 = ActiveWindow.Selection.ShapeRange(indexes(i))
                Set shp2 = ActiveWindow.Selection.ShapeRange(indexes(i + 1))

                ' 図形1の右端と図形2の左端を揃える
                shp2.left = shp1.left + shp1.Width
            Next i

        End If
    End If
End Sub

Sub AlignShapesVerticalStick()
    ' 図形をくっつけて並べる　縦
 
    ' 2個以上のShape選択中判定
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim numShapes%

            numShapes = ActiveWindow.Selection.ShapeRange.Count

        If numShapes >= 2 Then
            Dim shp1, shp2 As Shape
            Dim i%
            Dim tops() As Double
            Dim indexes() As Integer
            ReDim tops(1 To numShapes)
            ReDim indexes(1 To numShapes)
                        For i = 1 To numShapes
                tops(i) = ActiveWindow.Selection.ShapeRange(i).Top
                indexes(i) = i
            Next i

            InsertionSortIndex tops, indexes

            For i = 1 To numShapes - 1

                Set shp1 = ActiveWindow.Selection.ShapeRange(indexes(i))
                Set shp2 = ActiveWindow.Selection.ShapeRange(indexes(i + 1))

                                        ' 図形1の右端と図形2の左端を揃える
                shp2.Top = shp1.Top + shp1.Height
            Next i
        End If
    End If
End Sub

Sub InsertionSortIndex(vals() As Double, indexes() As Integer)
    Dim i&
    Dim j&
    Dim currentValue#
    Dim tmpIndex%

     For i = LBound(vals) + 1 To UBound(vals)
        currentValue = vals(i)
        j = i - 1
        tmpIndex = indexes(i)
        ' 適切な位置に要素を挿入する
        Do While j >= LBound(vals)
            If vals(j) > currentValue Then
                vals(j + 1) = vals(j)
                indexes(j + 1) = indexes(j)

            Else
                Exit Do
            End If
            j = j - 1
        Loop
        vals(j + 1) = currentValue
        indexes(j + 1) = tmpIndex
    Next i
End Sub



Sub ClipPath()
 Dim MyData As DataObject
 Set MyData = New DataObject
 
 MyData.SetText ActivePresentation.FullName
 MyData.PutInClipboard

End Sub




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






' 図形削除 & ペースト >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub DeleteAndPasteShape()
    On Error Resume Next

    ' 図形選択していたら削除
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ActiveWindow.Selection.ShapeRange.Delete
    End If

    ' コピーしている図形をペースト
    ActiveWindow.View.Paste

    Exit Sub
ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume Next
End Sub





' テキストボックス挿入 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub InsertNoWrapTextBox()
    ' 図形内で改行しないにチェックしたテキストボックスを挿入する
    ' あるいは、選択した図形の図形内で改行をしないにチェックをいれる

    ' 何も選択してない場合挿入する。余白0。
    If ActiveWindow.Selection.Type = ppSelectionNone Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Dim textbox As Shape

        Set textbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        ActiveWindow.Presentation.PageSetup.SlideWidth / 2, _
                        ActiveWindow.Presentation.PageSetup.SlideHeight / 2, 0, 0)
        textbox.TextFrame.DeleteText
        With textbox.TextFrame
            .MarginTop = 0
            .MarginRight = 0
            .MarginBottom = 0
            .MarginLeft = 0
        End With
        textbox.TextFrame.TextRange.Select
    End If
    ' テキスト選択中の場合、折り返ししないにチェックする
    If ActiveWindow.Selection.Type = ppSelectionText Then
        If ActiveWindow.Selection.TextRange.Parent.Parent.HasTextFrame Then
            ActiveWindow.Selection.TextRange.Parent.Parent.TextFrame2.WordWrap = msoFalse
        End If

    ' 1つ以上の図形を選択中の場合、すべての図形で折り返ししないにチェックする
    ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim selectedTextBox As Shape

        For Each selectedTextBox In ActiveWindow.Selection.ShapeRange
            If selectedTextBox.HasTextFrame Then
                selectedTextBox.TextFrame2.WordWrap = msoFalse
            End If
        Next selectedTextBox
    End If

    Exit Sub
ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume Next
End Sub








Sub ObjectsAlignTopLeft()
    ' 左上整列
    
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
            
    End If
    
    CommandBars.ExecuteMso "ObjectsAlignLeftSmart"
    CommandBars.ExecuteMso "ObjectsAlignTopSmart"
End Sub



Sub ObjectsAlignTopRight()
    ' 右上整列
    
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If
       
    CommandBars.ExecuteMso "ObjectsAlignRightSmart"
    CommandBars.ExecuteMso "ObjectsAlignTopSmart"
End Sub



Sub ShapeFillToNone()
    On Error GoTo ErrorHandler
    
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim selectedShape As Shape
        
        For Each selectedShape In ActiveWindow.Selection.ShapeRange
            selectedShape.Fill.Visible = msoFalse
        Next selectedShape
    End If
    
    Exit Sub
ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume Next
End Sub



Function isArrayEmpty(arr_var As Variant)
    Dim p As Integer
     
    On Error Resume Next
        p = UBound(arr_var, 1)
    If Err.Number = 0 Then
        isArrayEmpty = False
    Else
        isArrayEmpty = True
    End If
End Function



' 左上の位置に移動する >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub MoveToAnchor()
     ' 左上の位置に移動する

    Debug.Print ActiveWindow.Selection.Type; ppSelectionText
    
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        Exit Sub
    End If

    ActiveWindow.Selection.ShapeRange(1).left = 15.87402
    ActiveWindow.Selection.ShapeRange(1).Top = 60.52118

End Sub






' 表の幅を文字に合わせる       >>>> > > > > >> > > > > > > >> > > > >> > > >> > > >> >
Sub TableWidthAutoFit()

    If ActiveWindow.Selection.Type = ppSelectionNone Then Exit Sub
    If ActiveWindow.Selection.Type = ppSelectionSlides Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange(1).Type = msoTable Then Exit Sub
    
    ' テキストボックスを使って文字サイズをはかる。
    ' テキスト、フォント、文字サイズ、を合わせる。
    Dim i_col&, i_row&

    Dim table As table
    Debug.Print ActiveWindow.Selection.ShapeRange(1).Type

    Set table = ActiveWindow.Selection.ShapeRange(1).table

    Dim txtbox As Shape
    Dim horizontal_margin&
    Dim max_width&

    For i_col = 1 To table.Columns.Count
        max_width = 0
        For i_row = 1 To table.Rows.Count
            With table.Cell(i_row, i_col).Shape.TextFrame
                horizontal_margin = .MarginLeft + .MarginRight

            End With
            Set txtbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                            msoTextOrientationHorizontal, 0, 0, 0, 0)
            txtbox.TextFrame.WordWrap = msoFalse
            txtbox.TextFrame.Orientation = table.Cell(i_row, i_col).Shape.TextFrame.Orientation
            With txtbox.TextFrame.TextRange
                .Text = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Text
                With .Font
                    .Name = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Font.Name
                    .Bold = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Font.Bold
                    .Italic = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Font.Italic
                    .Size = table.Cell(i_row, i_col).Shape.TextFrame.TextRange.Font.Size
                End With

                If max_width < .BoundWidth + horizontal_margin Then max_width = .BoundWidth + horizontal_margin
            End With
            txtbox.Delete
            table.Columns(i_col).Width = max_width
        Next i_row
    Next i_col
End Sub




Sub test()
    TableWidthAutoFit
    
End Sub

Sub test1()
    Debug.Print ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Name
End Sub

Sub HandleError(ErrNumber As Long, ErrDescription As String)
    MsgBox "エラーが発生しました:" & vbCrLf & _
        "エラーコード: " & ErrNumber & vbCrLf & _
        "エラーメッセージ: " & ErrDescription, vbCritical, "エラー"
End Sub


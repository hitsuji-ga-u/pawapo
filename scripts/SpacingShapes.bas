Sub GetIntervalHorizontal(control As IRibbonControl, ByRef text)
    text = CStr(pt2cm(interval_horizontal))
End Sub

Sub GetIntervalVertical(control As IRibbonControl, ByRef text)
    text = CStr(pt2cm(interval_vertical))
End Sub

Sub SetIntervalHorizontal(control As IRibbonControl, ByRef text As String)
    if not isnumeric(text) Then
        text = CStr(interval_horizontal)
        ribbon.InvalidateControl("interval_horizontal")
        
        Exit Sub
    End If

    interval_horizontal = cm2pt(CDbl(text))
End Sub

Sub SetIntervalVertical(control As IRibbonControl, ByRef text As String)
    if not isnumeric(text) Then
        text = CStr(interval_vertical)
        ribbon.InvalidateControl("interval_vertical")
        Exit Sub
    End If

    interval_vertical = cm2pt(CDbl(text))
End Sub

' align shapes with no gaps between each other  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub SpacingShapesHorizontal()
    '  horizontaly align shapes with no gaps between each other

    ' only when selecting more than 1 shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Exit Sub
    End If

    Dim numShapes%

    numShapes = ActiveWindow.Selection.ShapeRange.Count
    If numShapes < 2 Then
        Exit Sub
    End If

    ' 回転後の図形の位置を取得
    Dim i As Integer
    Dim vertices() As Double
    Dim shps_left() As Double
    ReDim shps_left(1 To numShapes)

    For i = 1 To numShapes
        vertices = GetShapeConers(ActiveWindow.Selection.ShapeRange(i))
        shps_left(i) = min(vertices(0, 0), vertices(1, 0), vertices(2, 0), vertices(3, 0))
    Next i

    ' shps_leftの値で図形を並び替え
    Dim indexes_order() As Long
    Dim shps_sorted() As Shape
    ReDim shps_sorted(1 To numShapes)

    indexes_order = GetSortedIndexes(shps_left)
    For i = 1 To numShapes
        Set shps_sorted(i) = ActiveWindow.Selection.ShapeRange(indexes_order(i))
    Next i

    ' 基準図形を取得
    Dim idx_base As Integer
    Dim shp_base As Shape

    idx_base = FindIndex(indexes_order, numShapes)
    Set shp_base = shps_sorted(idx_base)

    ' 基準図形の左側の図形を並べる
    Dim shp As Shape
    Dim shp_prev As Shape
    dim diff as double
    dim vertices_prev() as double

    if idx_base > 1 Then
        Set shp_prev = shp_base
        For i = idx_base-1 To 1 Step -1
            Set shp = shps_sorted(i)
            vertices = GetShapeConers(shp)
            vertices_prev = GetShapeConers(shp_prev)
            diff = min(vertices_prev(0, 0), vertices_prev(1, 0), vertices_prev(2, 0), vertices_prev(3, 0)) _
                    - max(vertices(0, 0), vertices(1, 0), vertices(2, 0), vertices(3, 0)) - interval_horizontal
            shp.left = shp.left  + diff
            Set shp_prev = shp
        Next i
    End If

    ' 基準図形の右側の図形を並べる
    if idx_base < numShapes Then
        Set shp_prev = shp_base

        For i = idx_base+1 To numShapes
            Set shp = shps_sorted(i)
            vertices = GetShapeConers(shp)
            vertices_prev = GetShapeConers(shp_prev)
            diff = max(vertices_prev(0, 0), vertices_prev(1, 0), vertices_prev(2, 0), vertices_prev(3, 0)) _
                    - min(vertices(0, 0), vertices(1, 0), vertices(2, 0), vertices(3, 0)) + interval_horizontal
            shp.left = shp.left + diff
            Set shp_prev = shp
        Next i
    End If
End Sub

Sub SpacingShapesVertical()
    '  horizontaly align shapes with no gaps between each other

    ' only when selecting more than 1 shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Exit Sub
    End If

    Dim numShapes%

    numShapes = ActiveWindow.Selection.ShapeRange.Count
    If numShapes < 2 Then
        Exit Sub
    End If

    ' 回転後の図形の位置を取得
    Dim i As Integer
    Dim vertices() As Double
    Dim shps_top() As Double
    ReDim shps_top(1 To numShapes)

    For i = 1 To numShapes
        vertices = GetShapeConers(ActiveWindow.Selection.ShapeRange(i))
        shps_top(i) = min(vertices(0, 1), vertices(1, 1), vertices(2, 1), vertices(3, 1))
    Next i

    ' shps_topの値で図形を並び替え
    Dim indexes_order() As Long
    Dim shps_sorted() As Shape
    ReDim shps_sorted(1 To numShapes)

    indexes_order = GetSortedIndexes(shps_top)
    For i = 1 To numShapes
        Set shps_sorted(i) = ActiveWindow.Selection.ShapeRange(indexes_order(i))
    Next i

    ' 基準図形を取得
    Dim idx_base As Integer
    Dim shp_base As Shape

    idx_base = FindIndex(indexes_order, numShapes)
    Set shp_base = shps_sorted(idx_base)

    ' 基準図形の上側の図形を並べる
    Dim shp As Shape
    Dim shp_prev As Shape
    dim diff as double
    dim vertices_prev() as double

    if idx_base > 1 Then
        Set shp_prev = shp_base
        For i = idx_base-1 To 1 Step -1
            Set shp = shps_sorted(i)
            vertices = GetShapeConers(shp)
            vertices_prev = GetShapeConers(shp_prev)
            diff = min(vertices_prev(0, 1), vertices_prev(1, 1), vertices_prev(2, 1), vertices_prev(3, 1)) _
                    - max(vertices(0, 1), vertices(1, 1), vertices(2, 1), vertices(3, 1)) - interval_vertical
            shp.top = shp.top  + diff
            Set shp_prev = shp
        Next i
    End If

    ' 基準図形の下側の図形を並べる
    if idx_base < numShapes Then
        Set shp_prev = shp_base

        For i = idx_base+1 To numShapes
            Set shp = shps_sorted(i)
            vertices = GetShapeConers(shp)
            vertices_prev = GetShapeConers(shp_prev)
            diff = max(vertices_prev(0, 1), vertices_prev(1, 1), vertices_prev(2, 1), vertices_prev(3, 1)) _
                    - min(vertices(0, 1), vertices(1, 1), vertices(2, 1), vertices(3, 1)) + interval_vertical
            shp.top = shp.top + diff
            Set shp_prev = shp
        Next i
    End If
End Sub

' copy distance between shapes. >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 図形間の距離をコピー、ペースト >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub CopyShapeDistances()
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then Exit Sub
    
    Dim shp1 As Shape, shp2 As Shape
    
    With ActiveWindow.Selection
        If .ShapeRange(1).Top < .ShapeRange(2).Top Then
            Set shp1 = .ShapeRange(1)
            Set shp2 = .ShapeRange(2)
        Else
            Set shp1 = .ShapeRange(2)
            Set shp2 = .ShapeRange(1)
        End If
    End With

    ' ShapeDistanceY = shp2.Top - shp1.Top - shp1.Height
    interval_vertical = shp2.Top - shp1.Top - shp1.Height
    ribbon.InvalidateControl("interval_vertical")

    With ActiveWindow.Selection
        If .ShapeRange(1).left < .ShapeRange(2).left Then
            Set shp1 = .ShapeRange(1)
            Set shp2 = .ShapeRange(2)
        Else
            Set shp1 = .ShapeRange(2)
            Set shp2 = .ShapeRange(1)
        End If
    End With
    
    ' ShapeDistanceX = shp2.left - shp1.left - shp1.Width
    interval_horizontal = shp2.left - shp1.left - shp1.Width
    ribbon.InvalidateControl("interval_horizontal")
End Sub


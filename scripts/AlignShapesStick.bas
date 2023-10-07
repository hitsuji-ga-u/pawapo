
' align shapes with no gaps between each other  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub AlignShapesHorizontalStick()
    '  horizontaly align shapes with no gaps between each other

    ' only when selecting more than 1 shape
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

                shp2.left = shp1.left + shp1.Width
            Next i

        End If
    End If
End Sub

Sub AlignShapesVerticalStick()
    ' verticaly align shapes with no gaps between each other
 
    ' only when selecting more than 1 shape
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

                shp2.Top = shp1.Top + shp1.Height
            Next i
        End If
    End If
End Sub


' Add Nodes to square shape
Sub AddNodes()

    ' when not selecting shapes
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As shape

    For Each shp In ActiveWindow.Selection.ShapeRange
        If Not shp.Type = msoAutoShape Then GoTo continue

        If Not shp.AutoShapeType = msoShapeRectangle Then GoTo continue

        Dim x#, y#
        Dim i%, j%
        dim shpnd as ShapeNodes

        ' change to freeform
        with shp.nodes
            .Insert 1, msoSegmentLine, msoEditingAuto, shp.left, shp.top
            .Delete 2
        end with

        Dim shpVerticsCoordinate(1 To 8) As Double
        ' 図形の4つの角の座標を取得
        For i = 1 To 4
            shpVerticsCoordinate(i * 2 - 1) = shp.Nodes(i).Points(1,1)
            shpVerticsCoordinate(i * 2) = shp.Nodes(i).Points(1,2)
        Next i

        ' 中央の頂点を計算し、新しい頂点として追加
        For i = 1 To 4
            shp.Nodes.Insert i * 2 - 1 , msoSegmentLine, _
                msoEditingAuto, _
                (shpVerticsCoordinate(i * 2 - 1) + shpVerticsCoordinate(IIf(i = 4, 1, i + 1) * 2 - 1)) / 2, _
                (shpVerticsCoordinate(i * 2) + shpVerticsCoordinate(IIf(i = 4, 1, i + 1) * 2)) / 2
        Next i

continue:
    Next shp

End sub

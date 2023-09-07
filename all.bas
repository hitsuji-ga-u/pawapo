' Align Center >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub AlignCenterVertical()
    ' 1�߂ɑI�������}�`�̒����ɍ��킹��@�㉺����

    ' �}�`��I�����ĂȂ���ΏI���
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim shps As ShapeRange

    set shps = ActiveWindow.Selection.ShapeRange

    ' 1�̂ݑI���̏ꍇ
    If shps.Count = 1 Then
        shps.Align msoAlignMiddles, msoTrue

    ' 2�ȏ�I�����Ă���ꍇ
    Elseif shps.Count >= 2 Then
        Dim i&

        for i = 2 To shps.Count
            shps(i).Top = shps(1).Top + shps(1).Height/2 - shps(i).Height / 2
        next i
    end If
End sub

Sub AlignCenterHorizontal()
    ' 1�߂ɑI�������}�`�̒����ɍ��킹��@���E����

    ' �}�`��I�����ĂȂ���ΏI���
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim shps As ShapeRange

    set shps = ActiveWindow.Selection.ShapeRange

    ' 1�̂ݑI���̏ꍇ
    If shps.Count = 1 Then
        shps.Align msoAlignCenters, msoTrue

    ' 2�ȏ�I�����Ă���ꍇ
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



' �}�`���������ĕ��ׂ� >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub AlignShapesHorizontalStick()
    ' �}�`���������ĕ��ׂ�@��

    ' 2�ȏ��Shape�I�𒆔���
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

                ' �}�`1�̉E�[�Ɛ}�`2�̍��[�𑵂���
                shp2.left = shp1.left + shp1.Width
            Next i

        End If
    End If
End Sub

Sub AlignShapesVerticalStick()
    ' �}�`���������ĕ��ׂ�@�c
 
    ' 2�ȏ��Shape�I�𒆔���
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

                                        ' �}�`1�̉E�[�Ɛ}�`2�̍��[�𑵂���
                shp2.Top = shp1.Top + shp1.Height
            Next i
        End If
    End If
End Sub




' �}�`�̓h��Ԃ��F�A�g���̐F�A�t�H���g�̐F��ς��� > > > >> > > > >> > > > >> > > >

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
    ' �h��Ԃ��̐F��ς���
    ' color_idx: msoThemeColor RGB�Ŏw�肷��Ȃ�color_idx=0�ɂ���B
    ' -1�œh��Ԃ������B
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

    Else If ActiveWindow.selection.type = ppSelectionText then
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
    Else If ActiveWindow.selection.type = ppselectiontext then
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
    Else If ActiveWindow.selection.type = ppselectiontext then
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


' Clip Path >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub ClipPath()
 Dim MyData As DataObject
 Set MyData = New DataObject
 
 MyData.SetText ActivePresentation.FullName
 MyData.PutInClipboard

End Sub


' �}�`�Ԃ̋������R�s�[�A�y�[�X�g > > > > >> > > > >> > > >> > > >> > > >> > > >> > >> 
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

    ShapeDistanceY = shp2.Top - shp1.Top - shp1.Height


    With ActiveWindow.Selection
        If .ShapeRange(1).left < .ShapeRange(2).left Then
            Set shp1 = .ShapeRange(1)
            Set shp2 = .ShapeRange(2)
        Else
            Set shp1 = .ShapeRange(2)
            Set shp2 = .ShapeRange(1)
        End If
    End With
    
    ShapeDistanceX = shp2.left - shp1.left - shp1.Width
End Sub



' �}�`�Ԃ̋����y�[�X�g Y���� > > > > >> > > > >> > 

Sub PasteShpaeDistancesX()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then Exit Sub
    
    Dim i&
    
    For i = 2 To ActiveWindow.Selection.ShapeRange.Count
        With ActiveWindow.Selection
            .ShapeRange(i).Left = .ShapeRange(i - 1).Left + .ShapeRange(i - 1).Width + ShapeDistanceX
        End With
    Next i

End Sub

Sub PasteShpaeDistancesY()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then Exit Sub
    
    Dim i&
    
    For i = 2 To ActiveWindow.Selection.ShapeRange.Count
        With ActiveWindow.Selection
            .ShapeRange(i).Top = .ShapeRange(i - 1).Top + .ShapeRange(i - 1).Height + ShapeDistanceY
        End With
    Next i

End Sub
' �}�`�̈ʒu�R�s�[ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub CopyShapesPos()
    ' �}�`�̈ʒu���i�[����

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim selectedShapes As ShapeRange
    Dim i As Long

    Set selectedShapes = ActiveWindow.Selection.ShapeRange
    ReDim shapePositions(1 To selectedShapes.Count, 1 To 2) ' 2�����z�� (x, y)

    For i = 1 To selectedShapes.Count
        shapePositions(i, 1) = selectedShapes(i).left
        shapePositions(i, 2) = selectedShapes(i).Top
    Next i
End Sub

Sub PasteShapesAbsolutely()
    ' �I�������}�`���R�s�[���Ă���ʒu�ɐ�ΓI�ɍ��킹��

    ' �ʒu�R�s�[����ĂȂ���ΏI��
    If isArrayEmpty(shapePositions) Then
        Exit Sub
    End If

    ' �}�`�I������ĂȂ���ΏI��
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    Dim i&
    Dim selectedShpsNum As Integer

    ' �I�����ꂽ�}�`�̐����擾
    selectedShpsNum = ActiveWindow.Selection.ShapeRange.Count

    ' min(�}�`�̑I��, �R�s�[���Ă���}�`�̈ʒu��)�̐}�`�𒲐�����B
    For i = 1 To IIf(UBound(shapePositions) < selectedShpsNum, UBound(shapePositions), selectedShpsNum)
        With ActiveWindow.Selection.ShapeRange(i)
            .left = shapePositions(i, 1)
            .Top = shapePositions(i, 2)
        End With
    Next i
End Sub

Sub PasteShapesRelatively()
    ' �I�������}�`���R�s�[���Ă���ʒu�ɑ��ΓI�ɍ��킹��

    ' �ʒu�R�s�[����ĂȂ���ΏI��
    If isArrayEmpty(shapePositions) Then
        Exit Sub
    End If

    ' �}�`�I������ĂȂ���ΏI��
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then
        Exit Sub
    End If

    ' �ʒu�R�s�[����2�ȏ�Ȃ���ΏI��
    If UBound(shapePositions) - LBound(shapePositions) + 1 < 2 Then
        Exit Sub
    End If

    ' �}�`��ȏ�I������Ă��Ȃ���ΏI��
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        Exit Sub
    End If
        Dim i&

    Dim selectedShpsNum As Integer

    ' �I�����ꂽ�}�`�̐����擾
    selectedShpsNum = ActiveWindow.Selection.ShapeRange.Count

    ' min(�}�`�̑I��, �R�s�[���Ă���}�`�̈ʒu��)�̐}�`�𒲐�����B
    For i = 2 To IIf(UBound(shapePositions) < selectedShpsNum, UBound(shapePositions), selectedShpsNum)
        With ActiveWindow.Selection
            .ShapeRange(i).left = .ShapeRange(1).left + shapePositions(i, 1) - shapePositions(1, 1)
            .ShapeRange(i).Top = .ShapeRange(1).Top + shapePositions(i, 2) - shapePositions(1, 2)
        End With
    Next i
End Sub



' �}�`�폜 & �y�[�X�g >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub DeleteAndPasteShape()
    On Error Resume Next

    ' �}�`�I�����Ă�����폜
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ActiveWindow.Selection.ShapeRange.Delete
    End If

    ' �R�s�[���Ă���}�`���y�[�X�g
    ActiveWindow.View.Paste

    Exit Sub
ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume Next
End Sub
Sub DisableTextWrap()
    ' �}�`���ŉ��s���Ȃ��Ƀ`�F�b�N�����e�L�X�g�{�b�N�X��}������
    ' ���邢�́A�I�������}�`�̐}�`���ŉ��s�����Ȃ��Ƀ`�F�b�N�������

    On Error GoTo ErrorHandler

    ' �����I�����ĂȂ��ꍇ
    If ActiveWindow.Selection.Type = ppSelectionNone Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Dim textbox As Shape

        Set textbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        ActiveWindow.Presentation.PageSetup.SlideWidth / 2, _
                        ActiveWindow.Presentation.PageSetup.SlideHeight / 2, 0, 0)
        textbox.Select
        textbox.TextFrame.TextRange.Text = ""
    End If

    If ActiveWindow.Selection.Type = ppSelectionText Then
        If ActiveWindow.Selection.TextRange.Parent.Parent.HasTextFrame Then
            ActiveWindow.Selection.TextRange.Parent.Parent.TextFrame2.WordWrap = msoFalse
        End If

    ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim selectedTextBox As Shape

            Set selectedTextBox = ActiveWindow.Selection.ShapeRange(1)

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










Dim DrawExpandLines()


    
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange.Count = 2 Then Exit Sub

    If Not ActiveWindow.Selection.ShapeRange(1).Type = msoAutoShape And _
        Not ActiveWindow.Selection.ShapeRange(1).Type = msoPicture Then Exit Sub
    If Not ActiveWindow.Selection.ShapeRange(2).Type = msoAutoShape And _
        Not ActiveWindow.Selection.ShapeRange(2).Type = msoPicture Then Exit Sub

    Dim shp1 As Shape, shp2 As Shape
    Dim vertices() As Double
    Dim shp1a(1) As Double, shp1b(1) As Double, shp2a(1) As Double, shp2b(1) As Double
    Dim c1x#, c1y#, c2x#, c2y#
    
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    Set shp2 = ActiveWindow.Selection.ShapeRange(2)
    
    
    c1x = CDbl(shp1.left) + CDbl(shp1.Width) / 2
    c1y = CDbl(shp1.Top) + CDbl(shp1.Height) / 2
    c2x = CDbl(shp2.left) + CDbl(shp2.Width) / 2
    c2y = CDbl(shp2.Top) + CDbl(shp2.Height) / 2

    vertices = ShapeVertices(shp1)
    
    
    Dim i&, j&
    Dim bl_is_crossed As Boolean
    bl_is_crossed = False
    
    For i = 0 To 3
        j = (i + 1) Mod 4
        shp1a(0) = vertices(i, 0)
        shp1a(1) = vertices(i, 1)
        shp1b(0) = vertices(j, 0)
        shp1b(1) = vertices(j, 1)
        
        If is_crossed(shp1a(0), shp1a(1), shp1b(0), shp1b(1), c1x, c1y, c2x, c2y) Then
            bl_is_crossed = True
            Exit For
        End If
    Next i
    
    If Not bl_is_crossed Then Exit Sub
    Debug.Print "1"
    
    vertices = ShapeVertices(shp2)
    bl_is_crossed = False
    For i = 0 To 3
        j = (i + 1) Mod 4
        shp2a(0) = vertices(i, 0)
        shp2a(1) = vertices(i, 1)
        shp2b(0) = vertices(j, 0)
        shp2b(1) = vertices(j, 1)
        
        If is_crossed(shp2a(0), shp2a(1), shp2b(0), shp2b(1), c1x, c1y, c2x, c2y) Then
            bl_is_crossed = True
            Exit For
        End If
        
    Next i
    
    If Not bl_is_crossed Then Exit Sub
        Debug.Print "2"
       
    Dim ln1 As Shape
    Dim ln2 As Shape
    
    If is_crossed(shp1a(0), shp1a(1), shp2a(0), shp2a(1), shp1b(0), shp1b(1), shp2b(0), shp2b(1)) Then
        Set ln1 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1a(0), shp1a(1), shp2b(0), shp2b(1))
        Set ln2 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1b(0), shp1b(1), shp2a(0), shp2a(1))
    Else
        Set ln1 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1a(0), shp1a(1), shp2a(0), shp2a(1))
        Set ln2 = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex).Shapes.AddLine( _
                shp1b(0), shp1b(1), shp2b(0), shp2b(1))
    End If


End Sub




Function ShapeVertices(shp As Shape) As Variant
    '  ���v����4�p�̒��_���W�� �ԋp
    Dim vertices_0(3, 1) As Double
    Dim vertices(3, 1) As Double
    Dim cx#, cy#, s#, c#
    Dim i%
    
    Debug.Print "shp left, width = "; shp.Left, shp.Width
    Debug.Print TypeName(CDbl(shp.Width))
    
    Debug.Print "/2", CDbl(shp.Width) / 2
    Debug.Print "+", shp.Left + shp.Width / 2
    
    
    cx = CDbl(shp.Left) + CDbl(shp.Width) / 2
    cy = shp.Top + shp.Height / 2
    s = Sin(shp.Rotation * 3.14159265358979 / 180)
    c = Cos(shp.Rotation * 3.14159265358979 / 180)
        
    vertices_0(0, 0) = shp.Left - cx
    vertices_0(0, 1) = shp.Top - cy
    vertices_0(1, 0) = shp.Left + shp.Width - cx
    vertices_0(1, 1) = shp.Top - cy
    vertices_0(2, 0) = shp.Left + shp.Width - cx
    vertices_0(2, 1) = shp.Top + shp.Height - cy
    vertices_0(3, 0) = shp.Left - cx
    vertices_0(3, 1) = shp.Top + shp.Height - cy

    For i = 0 To 3
        vertices(i, 0) = vertices_0(i, 0) * c - vertices_0(i, 1) * s + cx
        vertices(i, 1) = vertices_0(i, 0) * s + vertices_0(i, 1) * c + cy
    Next

    ShapeVertices = vertices
End Function

Function is_crossed(Ax#, Ay#, Bx#, By#, Cx#, Cy#, Dx#, Dy#) As Boolean
    ' �_B, D�͋��E�����܂ނƂ���B

    Dim s#, t#

    s = (Cy - Ay) * (Bx - Ax) - (By - Ay) * (Cx - Ax)
    t = (Dy - Ay) * (Bx - Ax) - (By - Ay) * (Dx - Ax)

        If s * t > 0 Or s = 0 Then
        is_crossed = False
        Exit Function
    End If

    s = (Ay - Cy) * (Dx - Cx) - (Dy - Cy) * (Ax - Cx)
    t = (By - Cy) * (Dx - Cx) - (Dy - Cy) * (Bx - Cx)
        If s * t > 0 Or s = 0 Then
        is_crossed = False
        Exit Function
    End If

    is_crossed = True
End Function

Sub FrequentlyUse()

    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub

    Dim shp As Shape

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    Debug.Print shp.Type
    Debug.Print shp.AutoShapeType


    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = msoLine Or shp.AutoShapeType = msoShapeMixed Then
            
            shp.Line.EndArrowheadLength = msoArrowheadLong
            shp.Line.EndArrowheadWidth = msoArrowheadWide
            shp.Line.EndArrowheadStyle = msoArrowheadOpen
            shp.Line.Weight = 1
        End If
    Next

End Sub


' �e�L�X�g�{�b�N�X�}�� >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub InsertNoWrapTextBox()
    ' �}�`���ŉ��s���Ȃ��Ƀ`�F�b�N�����e�L�X�g�{�b�N�X��}������
    ' ���邢�́A�I�������}�`�̐}�`���ŉ��s�����Ȃ��Ƀ`�F�b�N�������

    ' �����I�����ĂȂ��ꍇ�}������B�]��0�B�܂�Ԃ����Ȃ��`�F�b�N�͈ȍ~�̏����ōs���B
    If ActiveWindow.Selection.Type = ppSelectionNone Or ActiveWindow.Selection.Type = ppSelectionSlides Then
        Dim textbox As Shape

        Set textbox = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox( _
                        msoTextOrientationHorizontal, _
                        ActiveWindow.Presentation.PageSetup.SlideWidth / 2, _
                        ActiveWindow.Presentation.PageSetup.SlideHeight / 2, 0, 0)

        textbox.TextFrame.DeleteText
        textbox.TextFrame.TextRange.Select
    End If

    ' �e�L�X�g�I�𒆂̏ꍇ�A�܂�Ԃ����Ȃ��Ƀ`�F�b�N����
    If ActiveWindow.Selection.Type = ppSelectionText Then
        If ActiveWindow.Selection.TextRange.Parent.Parent.HasTextFrame Then
            With ActiveWindow.Selection.TextRange.Parent.Parent.TextFrame
                .WordWrap = msoFalse
                .MarginTop = 0
                .MarginRight = 0
                .MarginBottom = 0
                .MarginLeft = 0
            End With
        End If

    ' 1�ȏ�̐}�`��I�𒆂̏ꍇ�A���ׂĂ̐}�`�Ő܂�Ԃ����Ȃ��Ƀ`�F�b�N����
    ElseIf ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim selectedTextBox As Shape

        For Each selectedTextBox In ActiveWindow.Selection.ShapeRange
            If selectedTextBox.HasTextFrame Then
                With selectedTextBox.TextFrame
                    .WordWrap = msoFalse
                    .MarginTop = 0
                    .MarginRight = 0
                    .MarginBottom = 0
                    .MarginLeft = 0
                End With
            End If
        Next selectedTextBox
    End If

    Exit Sub
ErrorHandler:
    HandleError Err.Number, Err.Description
    Resume Next
End Sub


' �}���\�[�g > > > > >> > > > > > > > > > > > > > > > >> > > >> >> > > > >
Sub InsertionSortIndex(vals() As Double, indexes() As Integer)
    ' Double�̔z��vars�̏����ŁAindexes����בւ���B
    Dim i&
    Dim j&
    Dim currentValue#
    Dim tmpIndex%

     For i = LBound(vals) + 1 To UBound(vals)
        currentValue = vals(i)
        j = i - 1
        tmpIndex = indexes(i)
        ' �K�؂Ȉʒu�ɗv�f��}������
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



' ����̈ʒu�Ɉړ����� >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub MoveToAnchor()
     ' ����̈ʒu�Ɉړ�����

    Debug.Print ActiveWindow.Selection.Type; ppSelectionText
    
    If Not (ActiveWindow.Selection.Type = ppSelectionShapes Or ActiveWindow.Selection.Type = ppSelectionText) Then
        Exit Sub
    End If

    ActiveWindow.Selection.ShapeRange(1).left = 15.87402
    ActiveWindow.Selection.ShapeRange(1).Top = 60.52118

End Sub



' �}�`�𔒂̃O���f�[�V�����ɂ���@ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub PaintGradation()

    Debug.Print ActiveWindow.Selection.Type; ppSelectionText

    if not ActiveWindow.selection.type = ppSelectionShapes then exit sub

    dim tgt_shp as Shape

    ' ���𖳂��ɂ���
    tgt_shp.Line.Visible = msoFalse

    ' �e�[�}�J���[��1�F�ڂ�h��Ԃ��Ɏg�p����
    shape.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
    shape.Fill.ForeColor.Brightness = 0

End Sub


' �\�̕��𕶎��ɍ��킹��       >>>> > > > > >> > > > > > > >> > > > >> > > >> > > >> >
Sub TableWidthAutoFit()

    If ActiveWindow.Selection.Type = ppSelectionNone Then Exit Sub
    If ActiveWindow.Selection.Type = ppSelectionSlides Then Exit Sub
    If not ActiveWindow.Selection.ShapeRange(1).Type = msoTable Then Exit Sub
    
    ' �e�L�X�g�{�b�N�X���g���ĕ����T�C�Y���͂���B
    ' �e�L�X�g�A�t�H���g�A�����T�C�Y�A�����킹��B
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
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub
    With ActiveWindow.Selection.ShapeRange(1)
        .Line.Visible = msoFalse
        With .Fill
            .ForeColor.ObjectThemeColor = msoThemeLight1
            .OneColorGradient msoGradientHorizontal, 3, 1
            .GradientStops(1).Transparency = 1
            .GradientStops(2).Position = 0.6
            .GradientStops(2).Transparency = 0.3
        End With
    End With
End Sub


' �����O���f�[�V����������  > > > > > > > > > > > > >> > > > > > > > > > > > > > > >
Sub TransGradation()
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then Exit Sub
    With ActiveWindow.Selection.ShapeRange(1)
        .Line.Visible = msoFalse
        With .Fill
            .ForeColor.ObjectThemeColor = msoThemeLight1
            .OneColorGradient msoGradientHorizontal, 3, 1
            .GradientStops(1).Transparency = 1
            .GradientStops(2).Position = 0.6
            .GradientStops(2).Transparency = 0.3
        End With
    End With
End Sub



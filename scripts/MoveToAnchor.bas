

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

Option Explicit
Dim shapePositions() As Variant
Dim ShapeDistanceX As Double
Dim ShapeDistanceY As Double
Dim margin_horizontal As Double
Dim margin_vertical As Double

Sub InitCustomTab()
    ShapeDistanceX = ActivePresentation.PageSetup.SlideWidth * 0.05
    ShapeDistanceY = ActivePresentation.PageSetup.SlideHeight * 0.01
    margin_horizontal = 0
    margin_vertical = 0
End Sub

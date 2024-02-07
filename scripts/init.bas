Option Explicit
Dim shapePositions() As Variant
Dim ShapeDistanceX As Double
Dim ShapeDistanceY As Double
Dim margin_horizontal As Double
Dim margin_vertical As Double
Dim total_page As Long
Dim ribbon As IRibbonUI
Dim edit_text As String



Sub InitCustomTab(rib As IRibbonUI)
    ShapeDistanceX = ActivePresentation.PageSetup.SlideWidth * 0.05
    ShapeDistanceY = ActivePresentation.PageSetup.SlideHeight * 0.01
    margin_horizontal = 0
    margin_vertical = 0

    ' ページ設定 初期化
    total_page = GetNowTotalPage

    Set ribbon = rib
End Sub

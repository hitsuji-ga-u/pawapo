
' 挿入ソート > > > > >> > > > > > > > > > > > > > > > >> > > >> >> > > > >
Sub InsertionSortIndex(vals() As Double, indexes() As Integer)
    ' Doubleの配列varsの昇順で、indexesを並べ替える。
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


Sub 給報に削除線を追加()

    Dim c As Range
    Dim p As Long
    Dim startPos As Long

    For Each c In Selection.Cells
        If Not c.HasFormula Then
            If VarType(c.Value2) = vbString Then
                startPos = 1
                Do
                    p = InStr(startPos, c.Value2, "給報")
                    If p = 0 Then Exit Do

                    c.Characters(p, 2).Font.Strikethrough = True
                    startPos = p + 2
                Loop
            End If
        End If
    Next c

End Sub

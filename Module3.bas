Option Explicit

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub my5_2()

    Rem my5_1と比べて行数は多くなるが、どのプロシージャがなんの処理を行っているか、スコープがわかりやすいためメンテナンス性も可読性も高くなった

    With Sheet1
        .Range("B1").Value = GetTaxIncluded(.Range("A1").Value)
        .Range("B2").Value = GetEndOfMonth(.Range("A2").Value)
    End With
    
End Sub

Function GetTaxIncluded(ByVal price As Long) As Currency

    Const TAX_RATE As Currency = 0.08
    GetTaxIncluded = price * (1 + TAX_RATE)
    
End Function

Function GetEndOfMonth(ByVal dt As Date) As Date

    GetEndOfMonth = DateSerial(Year(dt), Month(dt) + 1, 0)

End Function

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub my5_1()

    Const TAX_RATE As Currency = 0.08
    
    With Sheet1
        '税込価格を求める
        .Range("B1").Value = .Range("A1").Value * (1 + TAX_RATE)
        
        '月末日を求める
        Dim dt As Date: dt = .Range("A2").Value
        .Range("B2").Value = DateSerial(Year(dt), Month(dt), 1#)
    End With
    
End Sub


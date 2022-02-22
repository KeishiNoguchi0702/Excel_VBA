Option Explicit

Sub mysub02()
    
   '動的配列
   Dim numbers() As Long
   
   ReDim numbers(1 To 2) As Long
   numbers(1) = 10: numbers(2) = 30
   Debug.Print numbers(1), numbers(2)
   
   ReDim Preserve numbers(1 To 3)
   numbers(3) = 20
   Debug.Print numbers(1), numbers(2), numbers(3)
    
End Sub

Sub MySub()

    '固定配列の作成
    Dim numbers(1, 1 To 3) As Long
    numbers(0, 1) = 10: numbers(0, 2) = 20: numbers(0, 3) = 30
    numbers(1, 1) = 11: numbers(1, 2) = 21: numbers(1, 3) = 31
    
    Debug.Print numbers(0, 1), numbers(0, 2), numbers(0, 3)
    Debug.Print numbers(1, 1), numbers(1, 2), numbers(1, 3)
    
End Sub

Option Explicit

Private Type Person
    Name As String
    Age As Long
    Birthday As Date
End Type

Sub musub06()
    
    'ユーザ定義型
    Dim p As Person
    p.Name = "Bob"
    p.Age = 25
    p.Birthday = #1/1/1993#
    
    Debug.Print p.Name, p.Age, p.Birthday
    
End Sub

Sub mysub05()
 
    'Eraseキーワードで配列を初期化できる。ただし、動的配列は初期化後にメモリ解放するものの、固定配列は解放しない（ローカルウィンドウで確認）
   Dim staticArray(1 To 3) As Long
   staticArray(1) = 10: staticArray(2) = 30: staticArray(3) = 20
   
   Dim dinamicArray() As Long
   ReDim dinamicArray(1 To 3)
   dinamicArray(1) = 11: dinamicArray(2) = 31: dinamicArray(3) = 21
   
   Erase staticArray, dinamicArray
   Stop
    
End Sub

Sub mysub04()
    
    '配列に異なる型の値をまとめて格納
   Dim values As Variant
   values = Array("Bob", 25, #1/1/1993#)
   
   Dim i As Long
   For i = 0 To 2
        Debug.Print values(i)
   Next i
    
End Sub

Sub mysub03()
    
    'Variant型の変数と配列
   Dim values(1 To 3) As Variant
   values(1) = "Bob"
   values(2) = 25
   values(3) = #1/1/1993#
   
   Debug.Print values(1), values(2), values(3)
    
End Sub

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

Sub mysub()

    '固定配列の作成
    Dim numbers(1, 1 To 3) As Long
    numbers(0, 1) = 10: numbers(0, 2) = 20: numbers(0, 3) = 30
    numbers(1, 1) = 11: numbers(1, 2) = 21: numbers(1, 3) = 31
    
    Debug.Print numbers(0, 1), numbers(0, 2), numbers(0, 3)
    Debug.Print numbers(1, 1), numbers(1, 2), numbers(1, 3)
    
End Sub

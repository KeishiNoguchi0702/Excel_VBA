Option Explicit

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub my5_13()

    Rem Functionプロシージャの戻り値を受け取らずに、破棄するパターン
    Dim x As Long: x = 100
    GetTaxIncluded_2 x
    
End Sub

Function GetTaxIncluded_2(ByVal price As Long) As Currency

    Const TAX_RATE As Currency = 0.1
    GetTaxIncluded_2 = price * (1 + TAX_RATE)
    MsgBox GetTaxIncluded_2
    
End Function

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub my5_12()

    Dim x As Long, y As Long, z As Long
    x = 10: y = 30: z = 20
    
    Call Increment5_12(x, y, z)
    Debug.Print x, y, z
    
End Sub

Sub Increment5_12(ByVal x As Long, ParamArray num() As Variant)

    Rem 任意の数の引数を受け止めることができる。
    Rem 引数リストの最後のみに指定でき、自動的にオプションになる。上の例では、引数yとzが追加で格納されている
    
    Dim i As Long
    For i = LBound(num) To UBound(num)
        num(i) = num(i) + 1
    Next i
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------

Sub my5_11()

    Dim x(1 To 3) As Long
    x(1) = 10
    x(2) = 30
    x(3) = 20
    
    Call Increment_5_11(x)
    Debug.Print x(1), x(2), x(3)
    
End Sub

Sub Increment_5_11(ByRef num() As Long)

    Dim i As Long
    For i = LBound(num) To UBound(num)
        num(i) = num(i) + 1
    Next i
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------

Sub my5_9()

    Dim x As Long: x = 10
    Call Increment_x(x)
    Debug.Print x: Rem 引数を値渡ししているため、xは10のままで表示される
    
    Dim y As Long: y = 10
    Call Increment_y(y)
    Debug.Print y: Rem 引数を参照渡ししているため、呼び出し先のプロシージャで加算された結果であるy=11が表示される
    
End Sub

Sub Increment_x(ByVal num As Long)

    num = num + 1
    
End Sub

Sub Increment_y(ByRef num As Long)

    num = num + 1

End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------

Sub mysub5_8()

    Call Say_Proc("Hello", "Tom")
    Call Say_Proc("Goodbye")
    
End Sub

Sub Say_Proc(message As String, Optional name As String = "BoB")

    Rem 引数をオプション化することで、呼び出し元の引数の省略を可能にする
    MsgBox message & "," & name & "!"
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub5_7()
    
    Rem 順序による引数と名前付き引数の混在
    Call Say("Hello", name:="Bob")
    
End Sub

Sub mysyb5_6()

    Rem 名前付き引数:呼び出し先に渡す引数を名称で指定し､順序を入れ替えることができる
    Call Say(name:="Bob", message:="Hello")
    
End Sub

Sub mysub5_5()

    Call Say("Hello", "Bob")
    
End Sub

Sub Say(message As String, name As String)

    MsgBox message & "," & name & "!"
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub my5_3()

    Rem Callキーワードで呼び出す場合は、引数指定に()が必要(無いとエラーになる)
    Rem プロシージャを呼び出していることがわかりやすくなるように､Callキーワードは極力つけたほうがよい
    
    Call SayHello("Bob")
    
End Sub

Sub my5_4()
    
    Rem Callキーワードを付けない場合、()は不要
    SayHello "bob"
    
    Rem しかし､なくてもエラーにはならない
    Rem 理由は、四則演算を優先させるための()だと認識されるから。しかも、式の評価が可能で、強制的に値渡しになり実行できてしまうのでややこしい。
    Rem いずれにせよ、正しい挙動を予測しづらくなるため、Callをつけることを優先にする。何かしらの理由でCallなしにするならば、引数に()をつけない作法を一貫する。
    SayHello ("Bob")
    
End Sub

Sub SayHello(ByVal name As String)

    MsgBox "Hello," & name & "!"

End Sub


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


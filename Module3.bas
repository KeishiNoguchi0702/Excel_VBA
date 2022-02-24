Option Explicit

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


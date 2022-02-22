Option Explicit

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_31()

    Rem withステートメントは、基本的に1つのオブジェクトしか指定できないが、親子関係にあるオブジェクトはネスト構造によって指定可能
    
    With Sheet1.Range("A1")
        .Value = 1000
        .Interior.Color = RGB(255, 255, 0)
        With .Font
            .Name = "Meiryo UI"
            .Bold = True
            .Size = 8
        End With
    End With
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_30()

    Rem withステートメントを使用して、簡潔なコードで同じステートメントに繰り返しアクセス
    With Sheet1.Range("A1")
        .Value = 1000
        .Interior.Color = RGB(255, 255, 0)
        .Font.Bold = True
    End With
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_28()

    Rem Resumeステートメントでエラー処理ルーチンから処理を戻す
    On Error GoTo ErrorHandler
    
    Dim x As Long, y As Long
    x = 1
    Debug.Print x / y
    
    Exit Sub 'この記述がないとErrorHandler内を永遠にループしてしまう
    
ErrorHandler:
    y = 5
    Resume
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_27()

    Rem 有効になったエラーを無効にするOn Error GoTo 0ステートメントは、影響範囲を小さくするために積極的に使用したほうがよい
    Dim x As Long, y As Long
    x = 1
    
    On Error Resume Next
    Debug.Print x / y
    Debug.Print "エラーが無視されました"
    On Error GoTo 0
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_26()
    
    Rem 簡潔なコードになる一方、本当にエラーを無視してよいのか疑問。無視して良い革新がある場合を除き、使用しなことが無難。
    On Error Resume Next
    
    Dim x As Long, y As Long
    x = 1
    Debug.Print x / y
    Debug.Print "エラーが無視されました"
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub musub4_25()

    Rem On Error Go Toステートメントによるエラー処理
    On Error GoTo ErrorHandler
    
    Dim x As Long, y  As Long
    x = 1
    Debug.Print x / y
    
ErrorHandler:
    Debug.Print Err.Description
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_24()
    
    Rem 0で除算したためにエラーになるプロシージャ
    Dim x As Long, y As Long
    x = 1
    Debug.Print x / y
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_23()

    Rem Debug.Assertで異常値を検出する（本来0以上の整数しか入らないにもかかわらず、負数がセルに入っていた場合に処理を中断する）

    Dim cell As Range
    For Each cell In Sheet1.Range("A1:A10")
        Debug.Assert cell.Value >= 0
        Debug.Print cell.Value
    Next cell
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_22()

    Dim i As Long
    For i = 1 To 10
        Debug.Assert i Mod 5 <> 0 '条件式の結果がTRUEの場合は処理が中断されず、FALSEの場合に中断される
        Debug.Print "iの値：", i
    Next i
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_21()

    Rem iの値が５の倍数になるたびに一時停止する。ブレークポイントでは実現できない。Rialsのbinding.pryみたいなもの。
    Dim i As Long
    For i = 1 To 10
        If i Mod 5 = 0 Then Stop
        Debug.Print "iの値：", i
    Next i
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_20()

    Rem Gotoステートメントを使用することで、ループは継続させたまま、特定の条件をスキップ（あるいは実行）させることができる

    Dim i As Long
    
    For i = 1 To 10
        If i Mod 3 = 0 Then GoTo Continue
        Debug.Print "iの値：", i
Continue:
    Next i

End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_19()

    Rem Exitの対象をDoにしているため、外側の処理を終了させることができる
    Dim i As Long, j As Long
    i = 1
    
    Do While i <= 3
        For j = 1 To 3
            If i = 2 And j = 2 Then Exit Do
            Debug.Print i, j
        Next j
        i = i + 1
    Loop
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_18()

    Rem Exitステートメントで脱出するのは内側のForステートメントのみなので、外側のForステートメントの処理が継続する
    Dim i As Long, j As Long
    For i = 1 To 3
        For j = 1 To 3
            If i = 2 And j = 2 Then Exit For
            Debug.Print i, j
        Next j
    Next i
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_17()

    Dim y As Long
    Do
        y = Int(Rnd * 3) + 1
        Debug.Print "yの値：", y
        If y = 3 Then Exit Do
    Loop
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_16_1()

    Debug.Print "mysub4_16_2のcallステートメント前です"
    Call mysub4_16_2
    Debug.Print "mysub4_16_2のcallステートメント後です"
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_16_2()

    Debug.Print "mysub4_16_2のExitステートメント前です"
    Exit Sub
    Debug.Print "mysub4_16_2のExitステートメント後です"
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_15()

    Rem 後判定の場合は、直感的にuntilのほうがわかりやすい

    Rem whileキーワードによる反復制御
    Dim x As Long
    Do
        x = Int(Rnd * 3) + 1
        Debug.Print "xの値:", x
    Loop While x <> 3
    
    Dim y As Long
    Do
        y = Int(Rnd * 3) + 1
        Debug.Print "yの値：", y
    Loop Until x = 3
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_13()

    Rem どちらも同じ結果だけれど、whileのほうが直感的にわかりやすいかも

    Rem whileキーワードによる反復制御
    Dim x As Long: x = 1
    Do While x < 100
        Debug.Print "xの値：", x
        x = x * 3
    Loop
    
    Rem unitilキーワードによる反復制御
    Dim y As Long: y = 1
    Do Until y >= 100
        Debug.Print "yの値：", y
        y = y * 3
    Loop
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_12()

    Rem ネスト構造になっている配列でも、ループステートメントをネストしなくて済むのがとても大きなメリット
    Dim numbers(1, 1) As Long
    numbers(0, 0) = 0: numbers(1, 0) = 1
    numbers(0, 1) = 2: numbers(1, 1) = 3
    
    Dim number As Variant
    For Each number In numbers
        Debug.Print number
    Next number
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_11()

    Rem ubound関数とlbound関数で、配列の要素のインデックスを取得できる
    Dim numbers(1 To 3) As Long
    numbers(1) = 10
    numbers(2) = 30
    numbers(3) = 20
    
    Dim i As Long
    For i = LBound(numbers) To UBound(numbers)
        Debug.Print numbers(i)
    Next i
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_10()

    Rem 配列はlong型で宣言し、ループ処理のエレメントはVariant型で宣言する（しなければならない）
    Dim numbers(1 To 3) As Long
    numbers(1) = 10
    numbers(2) = 30
    numbers(3) = 20
    
    Dim number As Variant
    For Each number In numbers
        Debug.Print number
    Next number
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_9()

    Rem 可読性はmysub4_8に比べて劣るが、インデックスを付けてデバッグしたいときなどはこちらが有効

    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.Count
        Debug.Print i, ThisWorkbook.Sheets(i).Name
    Next i
    
    For i = 1 To Sheet1.Range("A1:C2").Count
        Debug.Print i, Sheet1.Range("A1:C2")(i).Address
    Next i
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_8()

    Rem for~eachステートメントは、集合の要素の数が不明でも取り扱える利点がある
    Rem 対象となるのは、コレクションまたは配列
    Rem 集合がコレクションの場合はVariant型またはObject型､配列の場合はVariant型で事前に宣言しなければならない
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        Debug.Print ws.Name
    Next ws
    
    Dim cell As Range
    
    For Each cell In Sheet1.Range("A2:C6")
        Debug.Print cell.Address
    Next cell
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_7()

    Rem forステートメント、増減両パターン
    Dim i As Long
    
    For i = 1 To 5
        Debug.Print "iの値:", i
    Next i
    
    For i = 10 To 0 Step -3
        Debug.Print "iの値：", i
    Next i
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_6()

    Rem case文に条件式を追加したいときは、select case trueを冒頭に各行に条件を記述する
    Dim text As String: text = "HOGFUGABA"
    Dim msg As String
    
    Select Case True
        Case text Like "*HOGE*"
            msg = "HOGEを含みます"
        Case text Like "*FUGA*"
            msg = "FUGAを含みます"
    End Select
    
    Debug.Print msg
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_5()

    Dim char As String: char = "A"
    Dim msg As String
    
    Select Case char
        Case "0" To "9"
            msg = "半角の数字です"
        Case "A" To "Z", "a" To "z"
            msg = "半角のアルファベットです"
        Case Else
            msg = "半角の数字でも半カウのアルファベットでもありません"
    End Select
    
    Debug.Print msg
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_4()
    
    Dim point As Long: point = 49
    Dim msg As String
    
    Select Case point
        Case 100
            msg = "満点ですね！"
        Case 97, 98, 99
            msg = "ほぼ満点ですね！"
        Case 80 To 96
            msg = "すごいですね！"
        Case Is >= 50
            msg = "頑張りましたね"
        Case Else
            msg = "次頑張りましょう"
    End Select
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub musub4_3()

    Dim rank As String: rank = "優"
    Dim msg As String
    
    Select Case rank
        Case "優"
            msg = "すごいですね！"
        Case "良"
            msg = "頑張りましたね！"
        Case "可"
            msg = "ギリギリでしたね"
        Case Else
            msg = "次に頑張りましょう"
    End Select
    
    Debug.Print msg
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_2()

    Dim num As Long: num = 53
    Dim digits As Long, msg As String
    
    If num < 10 Then
        digits = 1
    ElseIf num < 100 Then
        digits = 2
    End If
    
    If digits = 1 Then msg = "１桁です" Else msg = "２桁以上です"
    
    Debug.Print msg
    
End Sub

Rem ----------------------------------------------------------------------------------------------------------------------------------------------
Sub mysub4_1()

    Dim digits As Long: digits = 2
    Dim msg As String
    
    If digits = 1 Then msg = "1桁です" Else msg = "2桁です"
    Debug.Print msg
    
End Sub


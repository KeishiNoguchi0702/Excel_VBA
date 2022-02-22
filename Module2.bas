Option Explicit

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

Sub mysub4_1()

    Dim digits As Long: digits = 2
    Dim msg As String
    
    If digits = 1 Then msg = "1桁です" Else msg = "2桁です"
    Debug.Print msg
    
End Sub


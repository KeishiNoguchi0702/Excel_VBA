Option Explicit

Sub mysub4_1()

    Dim digits As Long: digits = 2
    Dim msg As String
    
    If digits = 1 Then msg = "1桁です" Else msg = "2桁です"
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


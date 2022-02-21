Attribute VB_Name = "Module1"
Option Explicit

Sub Sample001()

    Dim rng As Range
    Dim i As Long, v As Long, m As Long
    
    For i = 2 To Cells(Rows.Count, 3).End(xlUp).Row
        Set rng = Cells(i, 3).MergeArea
        If rng.Count > 1 Then
            rng.UnMerge
            v = rng(1)
            rng = Int(rng(1) / rng.Count)
            m = v - (rng(1) * rng.Count)
            rng.Resize(m) = rng(1) + 1
        End If
    Next

End Sub

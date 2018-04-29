Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Or Target.HasFormula Then Exit Sub
    On Error Resume Next
    If Not Intersect(Target, Range("A9:A1000")) Is Nothing Then
        Application.EnableEvents = False
        Target = UCase(Target)
        Application.EnableEvents = True
    End If
    If Not Intersect(Target, Range("B9:B1000")) Is Nothing Then
        Application.EnableEvents = False
        Target = Application.Proper(Target)
        Application.EnableEvents = True
    End If
    If Not Intersect(Target, Range("C9:G1000")) Is Nothing Then
        Application.EnableEvents = False
        Target = UCase(Target)
        Application.EnableEvents = True
    End If
    On Error GoTo 0
End Sub
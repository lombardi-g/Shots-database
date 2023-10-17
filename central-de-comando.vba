Private Sub Worksheet_Change(ByVal Target As Range)

' Reset cell D4 on changes to avoid confusion

    If Target.Address = "$J$4" Then
        Me.Range("E4").ClearContents
    End If

End Sub

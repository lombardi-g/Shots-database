Sub UndoLastInsertion()
    Dim lastRow As Long
    Dim wsCentral As Worksheet
    Dim wsDB_Afavor As Worksheet
    Dim wsDB_Sofr As Worksheet
    Dim emptyMessage As String
    Dim tbl As ListObject

    Set wsCentral = ThisWorkbook.Sheets("Central-de-comando")
    Set wsDB_Afavor = ThisWorkbook.Sheets("DB_Fin_Afavor")
    Set wsDB_Sofr = ThisWorkbook.Sheets("DB_Fin_Sofr")
    
    emptyMessage = "Último lançamento já foi desfeito"
    
    ' Check if values in A9:I9 are empty (indicating previous undo)
    If WorksheetFunction.CountA(wsCentral.Range("A9:I9")) = 0 Then
        MsgBox emptyMessage
        Exit Sub
    End If
    
    ' Find the last row in the respective database sheet based on the value in J4
    If wsCentral.Range("J4").Value = "A favor" Then
        lastRow = wsDB_Afavor.Cells(wsDB_Afavor.Rows.Count, "A").End(xlUp).Row
        
        If lastRow > 1 Then
            ' Clear contents of specific cells in the last row
            wsDB_Afavor.Range("A" & lastRow & ":I" & lastRow).ClearContents
            ' Resize the table to exclude the cleared row
            Set tbl = wsDB_Afavor.ListObjects(1)
            tbl.Resize tbl.Range.Resize(lastRow - tbl.HeaderRowRange.Row)
        End If
        
    ElseIf wsCentral.Range("J4").Value = "Contra" Then
        lastRow = wsDB_Sofr.Cells(wsDB_Sofr.Rows.Count, "A").End(xlUp).Row
        If lastRow > 1 Then
            ' Clear contents of specific cells in the last row
            wsDB_Sofr.Range("A" & lastRow & ":I" & lastRow).ClearContents
            ' Resize the table to exclude the cleared row
            Set tbl = wsDB_Sofr.ListObjects(1)
            tbl.Resize tbl.Range.Resize(lastRow - tbl.HeaderRowRange.Row)
        End If
    End If
    
    ' Clear values in A9:I9
    wsCentral.Range("A9:I9").ClearContents
  

End Sub

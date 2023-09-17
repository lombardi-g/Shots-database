Private Sub CommandButton1_Click()
    Dim wsDB As Worksheet
    Dim lastRow As Integer
    
    Set wsDB = ThisWorkbook.Sheets("DB_Fin_Afavor")
    
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row + 1
  
    wsDB.Cells(lastRow, 1).Value = CDate(TxtDate.Value)
    wsDB.Cells(lastRow, 2).Value = Me.ListBox1.Value
    wsDB.Cells(lastRow, 3).Value = Me.ListBox2.Value
    wsDB.Cells(lastRow, 4).Value = Me.ListBox3.Value
    wsDB.Cells(lastRow, 5).Value = Me.ListBox4.Value
    wsDB.Cells(lastRow, 6).Value = Me.ListBox5.Value
    wsDB.Cells(lastRow, 7).Value = Me.ListBox6.Value
    wsDB.Cells(lastRow, 8).Value = Me.ListBox7.Value
    wsDB.Cells(lastRow, 9).Value = Me.ListBox8.Value

End Sub

Private Sub UserForm_Initialize()
    Dim wsList As Worksheet
    Dim lastline As Integer
    Dim lastcolumn As Integer
    Dim i As Integer
    Dim j As Integer
    
    Set wsList = ThisWorkbook.Sheets("LIST")
    
    lastcolumn = wsList.Cells(1, wsList.Columns.Count).End(xlToLeft).Column
    
    For i = 2 To lastcolumn
    
        lastline = wsList.Cells(wsList.Rows.Count, i).End(xlUp).Row
    
        Select Case i
            Case 2
                For j = 2 To lastline
                    Me.ListBox1.AddItem wsList.Cells(j, i).Value
                Next j
            Case 3
                For j = 2 To lastline
                    Me.ListBox2.AddItem wsList.Cells(j, i).Value
                Next j
            Case 4
                For j = 2 To lastline
                    Me.ListBox3.AddItem wsList.Cells(j, i).Value
                Next j
            Case 5
                For j = 2 To lastline
                    Me.ListBox4.AddItem wsList.Cells(j, i).Value
                Next j
            Case 6
                For j = 2 To lastline
                    Me.ListBox5.AddItem wsList.Cells(j, i).Value
                Next j
            Case 7
                For j = 2 To lastline
                    Me.ListBox6.AddItem wsList.Cells(j, i).Value
                Next j
            Case 9
                For j = 2 To lastline
                    Me.ListBox7.AddItem wsList.Cells(j, i).Value
                Next j
            Case 10
                For j = 2 To lastline
                    Me.ListBox8.AddItem wsList.Cells(j, i).Value
                Next j
            
        End Select
    Next i
    
    Me.TxtDate.Value = Date
    
End Sub

Sub UpdateDatabaseWithVerification()
    Dim centralSheet As Worksheet
    Dim favorSheet As Worksheet
    Dim sofrSheet As Worksheet
    Dim i As Integer
    
    ' Set references to the sheets
    Set centralSheet = ThisWorkbook.Sheets("Central-de-comando")
    Set favorSheet = ThisWorkbook.Sheets("DB_Fin_Afavor")
    Set sofrSheet = ThisWorkbook.Sheets("DB_Fin_Sofr")
    
    ' Check the value in cell J4 to determine the database
    Dim databaseSheet As Worksheet
    If centralSheet.Range("$J$4").Value = "A favor" Then
        Set databaseSheet = favorSheet
    Else
        Set databaseSheet = sofrSheet

    End If
    
    ' Add a new row in the respective database
    i = databaseSheet.Cells(databaseSheet.Rows.Count, 1).End(xlUp).Row + 1
    databaseSheet.Cells(i, 1).Resize(1, 9).Value = centralSheet.Range("A4:I4").Value
    
    ' Update the verification row (A9:I9) in "Central-de-comando"
    centralSheet.Range("A9:I9").Value = centralSheet.Range("A4:I4").Value
    
    ' Replicate the value of J4 in B7 verification
    centralSheet.Range("$B$7").Value = centralSheet.Range("$J$4").Value
    
End Sub

Sub ProcessData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim jsonNames As String
    Dim jsonSources As String
    Dim names As String
    Dim sources As String
    
    ' Set the worksheet to work with
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with your sheet name
    
    ' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize the JSON data strings
    jsonNames = "["
    jsonSources = "["
    
    ' Loop through each row of data
    For i = 2 To lastRow
        Dim nameValue As String
        Dim sourceValue As String
        Dim levelValue As String
        nameValue = ws.Cells(i, 2).Value
        sourceValue = ws.Cells(i, 1).Value
        levelValue = ws.Cells(i, 6).Value
        
        If nameValue <> "" Then
            names = names & """" & nameValue & ""","
        End If
        If sourceValue <> "" Then
            sources = sources & """" & sourceValue & ""","
        End If
    Next i
    
    ' Remove the trailing commas
    If Len(names) > 0 Then
        names = Left(names, Len(names) - 1)
    End If
    If Len(sources) > 0 Then
        sources = Left(sources, Len(sources) - 1)
    End If
    
    ' Construct the JSON arrays
    jsonNames = jsonNames & names & "]"
    jsonSources = jsonSources & sources & "]"
    
    ' Write the JSON arrays to cells K1 and K2
    ws.Range("O1").Value = jsonNames
    ws.Range("O2").Value = jsonSources
End Sub
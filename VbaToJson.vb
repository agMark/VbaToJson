Function getJson(dataRange As Range)
    'This function will convert the selected data to a JSON text string
    'The first row of the selection must be column headers
    'All data is formatted as text...it's easy enough to run parseFloat if need be later
    
    Dim numCols As Integer
    Dim numRows As Integer
    Dim jsonString As String
    
    numCols = dataRange.Columns.Count
    numRows = dataRange.Rows.Count
    
    X = dataRange.Cells(3, 1).Value
    
    jsonString = "{'data':["
    X = dataRange.Cells(1, 8)
    For i = 2 To numRows
        jsonString = jsonString & "{"
        For j = 1 To numCols
            jsonString = jsonString & "'" & dataRange.Cells(1, j).Value & "':"
            jsonString = jsonString & "'" & dataRange.Cells(i, j).Value & "'"
            If j < numCols Then
                jsonString = jsonString & ","
            End If
        Next j
        If i < numRows Then
            jsonString = jsonString & "},"
        Else
            jsonString = jsonString & "}"
        End If
    Next i
    
    jsonString = jsonString & "]}"
    
    getJson = jsonString
End Function

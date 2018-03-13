Function getJson(dataRange As Range, writeToFile As Boolean)
    'This function will convert the selected data to a JSON text string
    'The first row of the selection must be column headers
    'All data is formatted as text...it's easy enough to run parseFloat if need be later
    
    Dim numCols As Integer
    Dim numRows As Integer
    Dim jsonString As String
    
    'quotation marks in the data can cause the JSON to be invalid
    Dim doubleQuoteReplace As String
    Dim singleQuoteReplace As String
    doubleQuoteReplace = "~%#" ' note: Chr(34) is a double quote
    singleQuoteReplace = "^*&" ' note: Chr(39) is a single quote
    Dim dataValue As String
    
    
    numCols = dataRange.Columns.Count
    numRows = dataRange.Rows.Count
    
    
    jsonString = "{" & doubleQuoteReplace & "data" & doubleQuoteReplace & ":["
    X = dataRange.Cells(1, 8)
    For i = 2 To numRows
        jsonString = jsonString & "{"
        For j = 1 To numCols
            dataValue = dataRange.Cells(i, j).Value
            'Get rid of double and single quotes
            dataValue = replace(dataValue, Chr(13), " ") 'change new lines to spaces
            dataValue = replace(dataValue, Chr(10), "") 'get rid of carriage returns
            dataValue = replace(dataValue, """", doubleQuoteReplace) 'double quotes
            dataValue = replace(dataValue, "'", singleQuoteReplace) 'single quotes
            
            
            jsonString = jsonString & doubleQuoteReplace & dataRange.Cells(1, j).Value & doubleQuoteReplace & ":"
            jsonString = jsonString & doubleQuoteReplace & dataValue & doubleQuoteReplace
            If j < numCols Then
                jsonString = jsonString & ","
            End If
        Next j
        If i < numRows Then
            jsonString = jsonString & "},"
        Else
            jsonString = jsonString & "}"
        End If
        'Debug.Print i
    Next i
    
    jsonString = jsonString & "]}"
    
    jsonString = replace(jsonString, doubleQuoteReplace, Chr(34))
    If Len(jsonString) > 32767 Then
        If writeToFile = True Then
            getJson = "To many characters to display in cell...trying to write data to jsonString.js"
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
        
            Dim Fileout As Object
            Set Fileout = fso.CreateTextFile("jsonString.js", True, True)

            Fileout.Write jsonString
            Fileout.Close
        Else
            getJson = "To many characters to display in cell...write to file disabled...see debug window"
        End If

    Else
        getJson = jsonString
        If writeToFile = True Then
            Dim fso2 As Object
            Set fso2 = CreateObject("Scripting.FileSystemObject")
        
            Dim Fileout2 As Object
            Set Fileout2 = fso2.CreateTextFile("jsonString.js", True, True)

            Fileout2.Write jsonString
            Fileout2.Close
        End If
    End If

End Function


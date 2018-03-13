# VbaToJson
Excel VBA Function to create JSON data

This was created to convert some data from a PDF -> Excel -> JSON for a project of mine.  This isn't really meant as an adaptable and re-usable function but is posted here so I don't lose it again.

This vba function will turn an excel range into a JSON string

In the selected range, the first line must be the column headers..
data is formatted as follows:
```javascript
{
  "data":[
    {
      "header1":"dataline1_cell1",
      "header2":"dataline1_cell2"
    },
    {
      "header1":"dataline2_cell1",
      "header2":"dataline2_cell2"    
    }
  ]
}
```
All data is formatted as strings so if you need numbers, parse with nodejs or something.

<h3>Usage:</h3> 
=getJson(A1:K39,TRUE)

First argument is the data range, Second argument is bool option to save to a file.  If true, JSON data is saved to jsonData.js in the same directory as the workbook.


# VbaToJson
Excel VBA Macro to create JSON file

This vba function will turn an excel array into a JSON string

In the selected range, the first line must be the column headers..
data is formatted as follows
{
  'data':[
    {
      'header1':'dataline1_cell1',
      'header2':'dataline1_cell2'
    },
    {
      'header1':'dataline2_cell1',
      'header2':'dataline2_cell2'    
    }
  ]
}

Usage: 
=getJson(A1:K39)



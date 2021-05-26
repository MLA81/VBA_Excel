# VBA_Excel

## Sheets / WorkSheets
``` vba
Dim sheet0 as Worksheet
sheet0 = Sheets(0)
sheet0 = ActiveSheet
```  

## Cells
``` vba
For row = 1 To row_Last
    If ActiveSheet.Cells(row, col_1).Value = "Wert" Then
        '' Do something
    End If
Next row
```

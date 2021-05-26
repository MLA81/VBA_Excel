# VBA_Excel

## Sheets / WorkSheets
``` vba
Dim sheet0 as Worksheet
sheet0 = Sheets(0)
sheet0 = ActiveSheet
```  

## Cells
``` vba
Dim row_Last As Long
row_Last = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).row

Dim col_last As Long
col_last = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
```

``` vba
For row = 1 To row_Last
    If ActiveSheet.Cells(row, col_1).Value = "Wert" Then
        '' Do something
    End If
Next row
```

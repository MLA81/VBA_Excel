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

## Dropdown mit Mehrfachauswahl

ℹ: Code ermöglicht es in der Range B2:B100 mehrere Werte in einem Dropdown auszuwählen
* XLSM erzeugen
* Code dem jeweiligen Tabellenblatt hinzufügen
* Evtl. Range anpassen - derzeit B2:B100
* Dropdown über Daten / Datentools / **Datenüberprüfung** hinzufügen

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
'** Mehrfachauswahl über DropDown-Liste (Gültigkeitsprüfung)
'** Einfügen im Code-Container des betreffenden Arbeitsblattes
 
'** Dimensionierung der Variablen
Dim rngDV As Range
Dim wert_old As String
Dim wertnew As String
 
'** Errorhandling
On Error GoTo Errorhandling
 
'** Mehrfachauswahl im definierten Bereich (Bsp. B2:B100) durchführen
If Not Application.Intersect(Target, Range("B2:B100")) Is Nothing Then
 
  '**Range definieren
  Set rngDV = Target.SpecialCells(xlCellTypeAllValidation)
  If rngDV Is Nothing Then GoTo Errorhandling
   
  '** Prüfen, ob eine gültige Zelle ausgewählt wurde und Werte eintragen
  If Not Application.Intersect(Target, rngDV) Is Nothing Then
    Application.EnableEvents = False
    wertnew = Target.Value
    Application.Undo
    wertold = Target.Value
    Target.Value = wertnew
    If wertold <> "" Then
      If wertnew <> "" Then
        Target.Value = wertold & ", " & wertnew
      End If
    End If
  End If
  Application.EnableEvents = True
End If
 
Errorhandling:
Application.EnableEvents = True
End Sub

```

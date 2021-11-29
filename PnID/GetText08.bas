Option Explicit
Sub ValveName()
'
' ValveName 매크로
'

'
    Range("F11").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(ValveLU,MATCH(LEFT(RC[-1],SEARCH(""-"",RC[-1])-1),ShortCode!R18C1:R25C1,0),2)"

    Range(Selection, Range("F" & Range("E11").End(xlDown).Row)).Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Option Explicit

Sub Block()
 Range("G11").Activate
 Dim no As Integer
  ActiveCell.FormulaR1C1 = _
        "=LEFT(SUBSTITUTE(RC[-2],LEFT(RC[-2],SEARCH(""-"",RC[-2])),""""),1)"
  no = ActiveCell.Value
  Do Until ActiveCell.Offset(0, -2) = ""
    If no = 1 Then
        ActiveCell = "1st Block"
     ElseIf no = 2 Then
        ActiveCell = "2nd Block"
     ElseIf no = 3 Then
        ActiveCell = "3rd Block"
     ElseIf no = 4 Then
        ActiveCell = "4th Block"
    ElseIf no = 5 Then
        ActiveCell = "5th Block"
    ElseIf no = 6 Then
        ActiveCell = "6th Block"
    ElseIf no = 7 Then
        ActiveCell = "7th Block"
    ElseIf no = 8 Then
        ActiveCell = "8th Block"
    ElseIf no = 9 Then
        ActiveCell = "9th Block"
    End If
    ActiveCell.Offset(1, 0).Activate
   Loop
End Sub

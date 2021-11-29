Option Explicit

Sub Valve_Man()
'
' Valve_Man 매크로
'

'
    Range("Summary").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "Valve_C"), CopyToRange:=Range("A10:D10"), Unique:=False
        

 ' 텍스트 숫자로 변경
 Dim rng As Range
 Dim ea As Range
 Dim i As Integer
 Set rng = Range(Range("C11"), Range("D11").End(xlDown))
     On Error Resume Next

    For Each ea In rng

        If IsNumeric(ea) Then

            ea = Format(ea, "#.####") '// 정상처리됨

        End If

    Next ea

 Set rng = Nothing
 
End Sub

Sub Valve_Data_Clear()
 Dim rng As Range
 Dim ea As Range
 Dim i As Integer
 Set rng = Range(Range("A11"), Range("D11").End(xlDown))
 rng.ClearContents
 

End Sub

Sub ValveTag정리()

Dim strTag As String
  Range("E11").Activate
     Do Until ActiveCell.Offset(0, -4) = ""
        strTag = ActiveCell.Offset(0, -4).Value
        ActiveCell = strTag
        ActiveCell.Offset(1, 0).Select
     Loop
  Range("E11").Select
End Sub

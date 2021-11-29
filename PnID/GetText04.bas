Option Explicit


Sub 도면찾기()
'
' PINO찾기 매크로
'
 Application.ScreenUpdating = False
 Dim shtPI As Worksheet
  Set shtPI = Worksheets("PI_NO")
 Dim shtLN As Worksheet
  Set shtLN = Worksheets("Line No. 정보 정리")
 Dim ipi As Integer
 Dim a As Integer
 Dim b As Integer
 Dim c As Integer
 Dim intNO As Integer
 shtPI.Select
  intNO = shtPI.Range(Range("A7"), Range("A7").End(xlDown)).Rows.Count + 6

 Dim intRow As Integer

 Dim rngPOX As Range
  Set rngPOX = shtPI.Range("I5")
  shtLN.Select
  intRow = shtLN.Range(Range("A10"), Range("A10").End(xlDown)).Rows.Count + 9
  shtPI.Select
 Dim rngPOY As Range
  Set rngPOY = shtPI.Range("J5")
 For b = 10 To intRow
  rngPOX = shtLN.Range("C" & b)
  rngPOY = shtLN.Range("D" & b)

 
  With shtPI
 
    For ipi = 7 To intNO
     .Cells(ipi, 9) = .Range("i5") >= .Cells(ipi, 4)
     .Cells(ipi, 10) = .Range("I5") <= .Cells(ipi, 5)
     .Cells(ipi, 11) = .Range("j5") <= .Cells(ipi, 6)
     .Cells(ipi, 12) = .Range("j5") >= .Cells(ipi, 7)
     .Cells(ipi, 13) = .Cells(ipi, 9) * .Cells(ipi, 10) * .Cells(ipi, 11) * .Cells(ipi, 12)
        For a = 7 To intNO
            If .Cells(a, 13) = 1 Then
                shtLN.Range("M" & b) = .Cells(a, 1)
        End If
        Next
    Next
   .Cells(5, 9).ClearContents
   .Cells(5, 10).ClearContents
   For c = 7 To intNO
    .Cells(c, 9).ClearContents
    .Cells(c, 10).ClearContents
    .Cells(c, 11).ClearContents
    .Cells(c, 12).ClearContents
    .Cells(c, 13).ClearContents
    
   Next
   End With

 Next
 shtLN.Select
 
End Sub

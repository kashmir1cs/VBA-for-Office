Option Explicit

Sub 도면넘버필터기준()
 Application.ScreenUpdating = False '화면전환 효과 비활성화
  With Cells(2, 14)
  .Value = "값"
  .Offset(1, 0) = "[*]"
  End With
End Sub

Sub 도면넘버추출()
'P&ID도면 넘버 좌표 정리
 Call 도면넘버필터기준
   
  Worksheets("PI_NO").Range(Range("A7"), Range("C7").End(xlDown)).ClearContents
    
  With Worksheets("PI_NO")
  
  Range("Summary").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "N2:N3"), CopyToRange:=Range("A6:C6"), Unique:=False
  End With
End Sub


Sub 도면번호정리()
 Application.ScreenUpdating = False '화면전환 효과 비활성화
 Worksheets("PI_NO").Select
 Call 도면넘버필터기준
 Call 도면넘버추출
   With Cells(2, 14)
  .Value = ""
  .Offset(1, 0) = ""
   End With
 Call 꼭지점위치
 '
' 테두리표시

 
  Dim rngAll As Range
  Dim intCol As Integer
  Set rngAll = Worksheets("PI_NO").Range(Range("A6"), Range("G6").End(xlDown))
    rngAll.Select
    
  With Selection
    Selection.Borders.LineStyle = xlContinuous
  End With
 ' 텍스트 숫자로 변경
 Dim rng As Range
 Dim ea As Range
 Dim i As Integer
 Set rng = Range(Range("B7"), Range("C7").End(xlDown))
     On Error Resume Next

    For Each ea In rng

        If IsNumeric(ea) Then

            ea = Format(ea, "#.####") '// 정상처리됨

        End If

    Next ea

 Set rng = Nothing
 
End Sub

Sub 꼭지점위치()
 Dim i As Integer
 Dim 넘버 As Integer
  넘버 = Range(Range("A7"), Range("A7").End(xlDown)).Rows.Count
  For i = 7 To 넘버 + 6
   Cells(i, 4) = Cells(i, 2) + Cells(3, 10)
   Cells(i, 5) = Cells(i, 2) + Cells(4, 10)
   Cells(i, 6) = Cells(i, 3) + Cells(3, 12)
   Cells(i, 7) = Cells(i, 3) + Cells(4, 12)
   Cells(i, 1).Replace what:="[", replacement:=""
   Cells(i, 1).Replace what:="]", replacement:=""
   
  Next
  
 End Sub
   
Sub 라인넘버중복삭제()
 Dim co As Long, i As Long
 Dim join As String
 co = Range(Range("C11"), Range("C11").End(xlDown)).Rows.Count
 
 Cells(11, 3).Select
 
 For i = 11 To co + 10
  join = Selection.Offset(0, 0) '오류 유형 선택'
   If join = "안내" Then
      Selection.EntireRow.Delete
      co = co - 1
   Else
      Selection.Offset(1, 0).Select
   End If
 Next
 
End Sub



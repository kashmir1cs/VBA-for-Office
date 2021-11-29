Option Explicit
Sub 정보추출()
 Dim str구분 As String '구분자 선언
     str구분 = "_"
 Dim int라인넘버 As Integer
 Dim i As Integer
  int라인넘버 = Range(Range("A10"), Range("A10").End(xlDown)).Rows.Count
  
'Line no.에서 정보 정리 실시
    For i = 10 To int라인넘버 + 9
     Cells(i, 5) = 문자분류(Cells(i, 1), 2, str구분) '유체추출
     Cells(i, 6) = 문자분류(Cells(i, 1), 1, str구분) 'Line Size 추출
     Cells(i, 7) = 문자분류(Cells(i, 1), 3, str구분) 'Serial No. 추출
     Cells(i, 8) = 문자분류(Cells(i, 1), 4, str구분) 'Line Spec. 추출
     Cells(i, 9) = 문자분류(Cells(i, 1), 5, str구분) 'Insultion 추출
          
    Next
     Call 설계압력
End Sub
Public Function 문자분류(str, n, sepChar)
' 언더바로 구분된 Line No.에서 각종 정보를 분리
 Dim f As Variant
  f = Split(str, sepChar)
  If n > 0 And n - 1 <= UBound(f) Then
    문자분류 = f(n - 1)
  Else
    문자분류 = "N/A"
  End If
  
End Function
Public Function Design압력(str)
' 설계 압력 계산 하는 함수 선언
 If str = "A" Then
    Design압력 = "10bar"
 ElseIf str = "B" Then
    Design압력 = "20bar"
 ElseIf str = "C" Then
    Design압력 = "30bar"
 ElseIf str = "AB" Then
    Design압력 = "15bar"
 ElseIf str = "" Then
    Design압력 = ""
 End If
 
End Function
Sub 설계압력()
 Dim int라인넘버 As Integer
 Dim i As Integer
  int라인넘버 = Range(Range("A10"), Range("A10").End(xlDown)).Rows.Count
  
'Line no.에서 정보 정리 실시
    For i = 10 To int라인넘버 + 9
     Cells(i, 11) = Design압력(Left(Cells(i, 8), 1)) '설계압력
          
    Next
End Sub

Sub valvecode()


End Sub

Sub 매크로2()
'
' 매크로2 매크로
'
  Application.ScreenUpdating = False
  
  
  Call criteria
   
    Range(Range("A9"), Range("D9").End(xlDown)).ClearContents
  
    Range("Summary").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "B1:C3"), CopyToRange:=Range("A9:D9"), Unique:=False
  
  Call datafont
    Call criteria_delete
    Call 언더바추가
    Call 테두리표시
    Call 정보추출
    Call 설계온도
    Call ChangeStrToInt
End Sub

Sub criteria()
 With Cells(1, 2)
  .Value = "하이픈"
  .Offset(0, 1) = "도면층"
  .Offset(1, 1) = "LINE NO PH-1"
  .Offset(1, 0) = "=LEN(RawData!A2)-LEN(SUBSTITUTE(RawData!A2,""-"",""""))>=3"
  .Offset(2, 1) = "UPW-E-Ph1"
  .Offset(2, 0) = "=LEN(RawData!A2)-LEN(SUBSTITUTE(RawData!A2,""-"",""""))>=3"

 End With
 
End Sub

Sub datafont()
 With Worksheets("Line No. 정보 정리")
  
  Cells.Font.Name = "맑은 고딕"
 End With
End Sub


Sub criteria_delete()
 Range("B1:C3").ClearContents
End Sub


Sub 언더바추가()
'
' 매크로5 매크로
'

'
    Range("A10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace what:="-", replacement:="_"
End Sub

Sub 테두리표시()
'
' 테두리표시
 
  Dim rngAll As Range
  Dim intCol As Integer
  Set rngAll = Range(Range("A9"), Range("D9").End(xlDown)).Resize(, 14)
    rngAll.Select
    
  With Selection
    Selection.Borders.LineStyle = xlContinuous
  End With
  
End Sub
Sub 정리()
' 시트에 입력된 내용 삭제
 Dim int라인넘버 As Integer
  int라인넘버 = Range(Range("A10"), Range("A10").End(xlDown)).Rows.Count + 9
 Range(Range("A10"), Range("M" & int라인넘버)).ClearContents
End Sub

Sub 설계온도()
' 유체별 온도 자동 입력
 Dim i As Integer
 Dim int라인넘버 As Integer
  int라인넘버 = Range(Range("A10"), Range("A10").End(xlDown)).Rows.Count
 Dim rngFluid As Range
 Set rngFluid = Range("FLUID")
 Dim rngFindex As Range
 Set rngFindex = Range("FINDEX")
 Worksheets("Line No. 정보 정리").Select
'Fluid에서 사용유체 온도값 추출
 With Application.WorksheetFunction
    For i = 10 To int라인넘버 + 9
     
     Cells(i, 12) = .Index(rngFluid, .Match(Cells(i, 5), rngFindex, 0), 3)
     
    Next
  End With
End Sub
Sub 온도구하기()
'
' 온도구하기 매크로
'
End Sub
Sub 자동필터설정()
'
' 자동필터설정
'

'
    Range("A9:N9").Select
    Selection.AutoFilter
    Range("A7").Select
End Sub

Sub ChangeStrToInt()
 Dim rng As Range
 Dim ea As Range
 Dim i As Integer
 Set rng = Range(Range("C10"), Range("D10").End(xlDown))
     On Error Resume Next

    For Each ea In rng

        If IsNumeric(ea) Then

            ea = Format(ea, "#.####") '// 정상처리됨

        End If

    Next ea

 Set rng = Nothing
 

End Sub

Sub ChangeDWGNOToInt()
 Dim rng As Range
 Dim ea As Range
 Dim i As Integer
 Set rng = Range(Range("M10"), Range("M10").End(xlDown))
     On Error Resume Next

    For Each ea In rng

        If IsNumeric(ea) Then

            ea = Format(ea, "#.####") '// 정상처리됨

        End If

    Next ea

 Set rng = Nothing
 

End Sub




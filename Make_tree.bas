Option Explicit
Dim lngMax As Long 'WBS 최대 Lv 숫자를 구함
'WBS Lv을 확인하여 Tree 구조로 변환
'status bar는 별도 form으로 작성필요 

Sub MakeTree()

Dim i As Integer


 상태창소환
 상태 ("데이터 포맷 변경")
 상태 ("개요수준 숫자로 변경")
 Text_to_Num_Level
 상태 ("완료")
 상태 ("ID 숫자로 변경")
 Text_to_Num_ID
 상태 ("완료")
 상태 ("기간 Data 숫자로 변경")
 기간to숫자
 상태 ("완료")
 
 'WBS Max Level 구하기
 findmaxlevel
 'lngMax (WBS Lv)만큼 Field 생성
  상태 ("기간 Data 숫자로 변경")
 For i = 1 To lngMax
 
    Range("L1").Offset(0, i).Value = "LV" & i
 
 Next
  Range("L1").Offset(0, lngMax + 1).Value = "Remark"
 Application.ScreenUpdating = False
'Activity 분리 시작
 상태 ("LV별 Tree구조로 정리")
 Activity정리
 상태 ("완료")
 상태 ("Tree구조화 작업 마무리")
 마무리
 상태 ("완료")
Pbar.Hide
 
 
End Sub

Sub findmaxlevel()
 Dim c As Range
 
 
 Set c = Range("I2", Range("I2").End(xlDown))
    lngMax = WorksheetFunction.Max(c)
 Set c = Nothing
    

End Sub
Sub Text_to_Num_Level()
 ' 개요수준 숫자로 변경
 Dim c As Range
 
 Set c = Range("I2", Range("I2").End(xlDown))
    c.EntireColumn.TextToColumns
 Set c = Nothing

End Sub

Sub Text_to_Num_ID()
 ' ID 숫자로 변경
 Dim c As Range
 
 Set c = Range("A2", Range("A2").End(xlDown))

    c.EntireColumn.TextToColumns
 Set c = Nothing

End Sub

Sub 기간to숫자()
'기간에서 숫자만 추출

Dim rng As Range
Dim r As Range
  Set rng = Range("E2", Range("E2").End(xlDown))
'반복문이용한 숫자만 추출
For Each r In rng
    상태 (r.Address)
    r = Replace(r, " 일", "")
    r = Replace(r, Chr(63), "") '물음표 제거
 

Next


End Sub



Sub 작업_테이블시트선택()

'원본시트선택
ActiveWorkbook.Sheets("작업_테이블").Activate

End Sub


Sub 원본데이터숨기기()
'Data 정제 후 Raw Data는 숨김
Range("E:J").EntireColumn.Hidden = True

End Sub
'// Status bar는 별도 form 작성 필요
Sub 상태창소환()

Pbar.Label1 = "Data 정리 작업 시작"
'Progress Bar 위치 설정
Pbar.Top = Application.UsableHeight / 2 + Pbar.Height
Pbar.Left = Application.UsableWidth / 2 - Pbar.Width / 2
    
Pbar.Show

End Sub
'// Status bar는 별도 form 작성 필요
Sub 상태(strMsg As String)

Pbar.Label1 = strMsg
Pbar.Repaint
End Sub
'// 좌측하단 상태창 표시 
Sub 하부창(strMsg As String)

Application.StatusBar = strMsg
End Sub

Sub Activity정리()

'test위한 procedure 생성
Dim r As Range
Dim strTemp As String
Dim intLv As Integer
Dim intCol As Integer
Dim intRow As Integer
Dim rngWrite As Range
Dim rngAct As Range
Dim intOffset As Integer
Dim rngLast As Range
 Set rngLast = Range("L1048576")


  Set rngAct = Range("D2", Range("D2").End(xlDown))

'셀마지막행 번호 1048576

'반복문 실행
'Tree구조로 
For Each r In rngAct
     상태창소환
intOffset = r.Offset(0, 5).Value 'wbs lv 값을 읽음
    
    상태 ("Activity 정리 :" & "Lv" & intOffset & " - " & r.Value)
    If r.Row = 2 Then
        
        Cells(2, 12 + intOffset) = r.Value
    
    '3행이후 데이터는 케이스별로 분류하여 정리
    '하위 레벨인 경우 옆에다 표기
    '윗 행 보다 lv가 높을 경우 계속해서 현재 행 옆 셀에 activity 표기
    ElseIf r.Row > 2 And r.Offset(0, 5).Value > 1 And r.Offset(-1, 5).Value < r.Offset(0, 5).Value Then

      intRow = Range("M1").CurrentRegion.Rows.Count '전체 영역의 Row Count하여 행의 현재 위치 지정 
      intCol = Cells(2, 12 + intOffset).Column  'wbs lv만큼 떨어진 위치 지정

      Cells(intRow, intCol) = r.Value 
    
    ElseIf r.Row > 2 And r.Offset(0, 5).Value > 1 And r.Offset(-1, 5).Value >= r.Offset(0, 5).Value Then
     
      intRow = Range("M1").CurrentRegion.Rows.Count + 1 '현재 range 값이 2보다 크고 윗 행 값보다 작거나 같은 경우 행을 바꿈
      intCol = Cells(2, 12 + intOffset).Column  'wbs lv만큼 떨어진 위치 지정
     
      Cells(intRow, intCol) = r.Value
      '범위 지정하여 상위 wbs lv 값을 표시
      'wbs lv 값을 읽어서 정해진 개수의 셀만큼 범위 지정하여 상위 LV wbs activity를 Cell에 입력
      Range(Cells(intRow, intCol - 1), Cells(intRow, intCol - intOffset)).Value = Range(Cells(intRow, intCol - 1), Cells(intRow, intCol - intOffset)).Offset(-1, 0).Value
      '현재 range의 WBS가 lv1일 경우 LV1에 해당 내용 표시
    ElseIf r.Row > 2 And r.Offset(0, 5).Value = 1 Then
     
      rngLast.Offset(0, 1).End(xlUp).Offset(1, 0) = r.Value
    
    End If
Next
End Sub

Sub 마무리()
'Tree 구조 최종 정리
'하위 Lv없는 Activity는 "-"라고 표시

Dim intCol As Integer
intCol = Range("M1").CurrentRegion.Columns.Count - 2
Range("M1", Range("M1").End(xlDown).Offset(0, intCol)).Select
Selection.SpecialCells(xlCellTypeBlanks).Select '값이 없는 Cell 선택
Selection.FormulaR1C1 = "-" '선택된 셀에 "-" 일괄 입력
End Sub

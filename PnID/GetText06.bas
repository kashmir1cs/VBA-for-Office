Option Explicit

Sub 정렬()
'
' 정렬 매크로
' 라인넘버/도면번호 기준으로 정렬
'

'
    Range("LINENO_DATA").Select
    ActiveWorkbook.Worksheets("Line No. 정보 정리").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Line No. 정보 정리").Sort.SortFields.Add Key:=Range( _
        Range("M10"), Range("M10").End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Line No. 정보 정리").Sort.SortFields.Add Key:=Range( _
        Range("A10"), Range("A10").End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Line No. 정보 정리").Sort
        .SetRange Range("LINENO_DATA")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub 라인넘버정렬()
 Call ChangeDWGNOToInt
 Call 정렬
End Sub

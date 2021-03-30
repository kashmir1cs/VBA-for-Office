Attribute VB_Name = "Module1"
Sub tstProgressBar()

' 먼저 진행상태를 표시할 폼을 보이게 합니다.
' 진도를 나타낼 Lable2의 너비는 0에서 출발합니다.
' 그리고 Lable3에서는 총 작업수와 현재 처리자료를 문자로 나타냅니다.
UserForm1.Label2.Width = 0
UserForm1.Label3 = "0 / 1000"
UserForm1.Show

' 처리해야 할 작업은 1000개로 정해진 상황으로 가정합니다.
For i = 1 To 1000
    ' 아래에 처리하는 작업이 위치합니다.
    ' 본 예제에서는 단순한 루프를 삽입하여 이를 대신합니다.
    For j = 1 To 1000
    ' 본 예제에서 처리 지연은 아래 루프의 수를 조절하여 정할 수 있습니다.
        For k = 1 To 1000
        Next k
    Next j
    
    ' 이어서 진행상태를 표시할 폼을 수정합니다.
    ' 기억해 두었던 레이블 너비 414가 여기서 사용됩니다.
    UserForm1.Label2.Width = Int(i / 1000 * 414)
    UserForm1.Label3 = Trim(i) + " / 1000"
    UserForm1.Repaint
Next i

' 모든 작업이 끝나면 폼을 숨깁니다.
UserForm1.Hide
    
End Sub

Sub 도면화()

Dim 선택 As Byte

     선택 = MsgBox("도면작성을 하시겠습니까?", vbYesNoCancel + vbQuestion, "오리피스 도면화")

 If 선택 = vbYes And Range("E15") = "Single Orifice" Then
    
    Call Orifice_1단_DWG_Data입력_도면생성
 ElseIf 선택 = vbYes And Range("E15") = "Double Orifice" Then
    Call Orifice_2단_DWG_Data입력_도면생성
 ElseIf 선택 = vbYes And Range("E15") = "Tripple Orifice" Then
    Call Orifice_3단_DWG_Data입력_도면생성
 ElseIf 선택 = vbNo Then
    Exit Sub
 End If
End Sub

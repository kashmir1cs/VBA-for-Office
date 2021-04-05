Sub 이름보이게하기()
Dim n As Name

On Error Resume Next
For Each n In ThisWorkbook.Names
n.Visible = True
Next
End Sub
Sub ShtCopy_1()
 Worksheets("도면_Single").Copy after:=Worksheets(1)
End Sub


Sub ShtCopy_2()
 Worksheets("도면_Double").Copy after:=Worksheets(1)

End Sub

Sub ShtCopy_3()
 Worksheets("도면_Tripple").Copy after:=Worksheets(1)

End Sub


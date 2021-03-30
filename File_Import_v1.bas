Attribute VB_Name = "Module1"
Option Explicit

Sub WeldPlanMerger_v13()
Form.Show


End Sub

Sub lstToCollection(lst As MSForms.ListBox, col As Collection)
'// List Box의 목록을 Collection 개체로 변환
Dim i As Integer
    
    For i = 0 To lst.ListCount - 1
        col.Add lst.List(i, 0)
    
    Next
    


End Sub
Sub StatusLabel(lbl As MSForms.Label, strText As String)

lbl.Caption = strText

End Sub


Sub Listmove(lstOriginal As MSForms.ListBox, lstMove As MSForms.ListBox, Optional All As Boolean = False) '// list간 항목 이동하는 Procedure 정립
Dim i As Integer
Dim lstMovecount(), count As Integer

If All = True Then

    For i = 0 To lstOriginal.ListCount - 1
        
        lstMove.AddItem lstOriginal.List(i, 0)
    
    Next i
    
    lstOriginal.Clear

Else
    
    For i = 0 To lstOriginal.ListCount - 1
        
        If lstOriginal.Selected(i) = True Then
            
            lstMove.AddItem lstOriginal.List(i, 0)
            lstOriginal.Selected(i) = False
            ReDim Preserve lstMovecount(count)
            lstMovecount(count) = i
            count = count + 1
        End If
    Next
    If count > 0 Then
    For i = UBound(lstMovecount) To 0 Step -1
        
        lstOriginal.RemoveItem lstMovecount(i)
        
    Next
    End If

End If

End Sub
Public Function splice(str, n, sepChar)
' 언더바로 구분된 Line No.에서 각종 정보를 분리
 Dim f As Variant
  f = Split(str, sepChar)
  If n > 0 And n - 1 <= UBound(f) Then
    splice = f(n - 1)
  Else
    splice = ""
  End If
  
End Function

Function TextCount(Text, Search As String) As Integer
   'Text에 들어있는 Search의 개수
   Dim i As Integer
   TextCount = 0
   For i = 1 To Len(Text)
      If Mid(Text, i, 1) = Search Then TextCount = TextCount + 1
   Next
End Function

Sub WeldPlanMerge(wbLists As Collection, Optional optSet As String = "Default")
'// 각 Workbook를 열면서 Range별로 복사/붙여넣기 반복
Dim i As Integer '// Progress Bar Width 조절용 변수

Dim no As Integer '// list 수량 Count
no = wbLists.count
Dim NoRow As Integer
Dim strFileName As String
Dim strDwg As String
Dim noSep As Integer '// "\" 개수 count
Dim wbList As Variant
Dim wb As Workbook
Dim wbTemp As Workbook
Dim startTime As Single
startTime = Timer
Dim finishTime As Single

'// Set wbTemp = Workbooks.Open("****.xlsx", False, False)
'//range 개체 선언

'// 각 WeldPlan 엑셀 파일의 Column을 Range 개체로 선언



Dim rng As Range '// Range Copy
Dim c As Range '// 반복문 적용 위한 Range 개체 선언

'// 범위 지정용 변수 선언

Dim endRow As Integer

Dim NoRows As Integer



'// Directory 현재로 변경

ChDir ThisWorkbook.Path & "\" '//현재 Directory로 변경

On Error Resume Next
Application.ScreenUpdating = False
Application.StatusBar = "준비"

If wbLists.count = 0 Then
    Exit Sub
    
Else
    If optSet = "Default" Then '// Setting이 "Default" 일 경우
        i = 1
        For Each wbList In wbLists
        '// WorkBook 개체 Wb에 할당
        Set wb = Workbooks.Open(wbList, False, True)
        noSep = TextCount(wbList, "\") '// directory Seperater 개수 Count
        strFileName = splice(wbList, noSep + 1, "\")
        Application.StatusBar = strFileName & " - File 열기 " & "(" & i & "of" & no & ")"
        On Error Resume Next
        Application.ScreenUpdating = False
            NoRow = Range("A2").End(xlDown).Row
            NoRows = NoRow - 2 + 1
            '// Workbook에서 각 Column 변수에 할당
            Application.StatusBar = strFileName & " - Data 읽기 " & "(" & i & "of" & no & ")"
           
            '//Template 파일 활성화
            wbTemp.Activate
            '// 범위 설정
            Set rng = wbTemp.Sheets(1).Range(Range("H1048576").End(xlUp).Offset(1, 0), Range("H1048576").End(xlUp).Offset(NoRows, 0))
            Application.StatusBar = strFileName & " - Data 복사하기 " & "(" & i & "of" & no & ")"
            '// JOINT 복사
            rngJoint.Copy rng
            '// RAW_MATERIAL 복사
            rngMatl.Copy rng.Offset(0, 1)

       
             Application.StatusBar = strFileName & " - FMCS용 Data 입력 " & "(" & i & "of" & no & ")"
            '// 해당하는 위치에 복사/붙여넣기 실시 
            '// Line입력
            
            'Status Bar 사용 
            Application.StatusBar = strFileName & " - 입력 완료 " & "(" & i & "of" & no & ")"
            
            wb.Close False
            i = i + 1
            Application.Wait (Now + TimeValue("0:00:01"))
            Next
        End If
  
    
    
End If

'//작업 완료 후 Message 출력 및 종료
Application.StatusBar = False
finishTime = Timer
MsgBox "총 " & no & "개 Merge 완료" & Chr(13) & Chr(10) & "소요시간 : " & finishTime - startTime & "sec"
wbTemp.SaveAs "C:\Export" & "\" & strSaveAs
End Sub



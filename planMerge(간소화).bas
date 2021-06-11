Option Explicit

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


Sub WeldPlanMerge()
'// File 선택창에서 선택하면 파일목록 저장할 List 선언
Dim WeldLists As Variant, Files As Variant

'// WorkBook 변수 선언
Dim wb As Workbook
'// Integer 변수 선언
Dim i As Integer
Dim NoRow As Integer
'range 개체 선언

Dim rngWP As Range


'Integer변수 선언

Dim endRow As Integer
Dim NoRows As Integer
Dim NoSep As Integer

'string 변수 선언
'Cell에 동일한 값을 붙여넣기 위한 변수
'File Name에서 추출
Dim FileName As String

'복사및 붙여넣기할 변수 선언
Dim rng As Range

ChDir ThisWorkbook.Path & "\"
WeldLists = Application.GetOpenFilename("Weld Plan Report, *.xlsx", Title:="Weld Plan 선택", MultiSelect:=True)
On Error Resume Next
Application.ScreenUpdating = False

If WeldLists <> False Then Else Exit Sub
For Each Files In WeldLists
    If InStr(Files, "__WP__") Then
        Debug.Print "파일처리중: " + Files
        NoSep = TextCount(Files, "\")
        Debug.Print NoSep
        '//파일명 추출하기
        FileName = splice(Files, NoSep + 1, "\")
        
        FileName = splice(Files, 4, "__")
        
        Debug.Print FileName
        
        Set wb = Workbooks.Open(Files, False, True)
        
        NoRow = Range("A2").End(xlDown).Row
        
        
        NoRows = NoRow - 2 + 1
        
        Debug.Print NoRows
        
        Set rngWP = wb.Sheets(1).Range(Range("A2"), Range("M" & Format(NoRow)))
            

        
        ThisWorkbook.Activate
    
        Set rng = ThisWorkbook.Sheets(1).Range(Range("A1048576").End(xlUp).Offset(1, 0), Range("A1048576").End(xlUp).Offset(NoRows, 0))
        '// Joint no.복사
        rngWP.Copy rng.Offset(0, 4)
        
        '// PCF 파일명 붙여넣기
        rng.Value = FileName
        '// JOBNO 붙여넣기
        '// String 편집하여 넣는 Field는 뒤쪽에 넣을 것
        rng.Offset(0, 1).Value = splice(FileName, 1, "-")
        '// Fluid 붙여넣기
        rng.Offset(0, 2).Value = splice(FileName, 5, "-")
        '// Line No. 붙여넣기
        rng.Offset(0, 3).Value = splice(FileName, 5, "-") & "-" & splice(FileName, 6, "-")
        wb.Close False
        
        
    End If
    
Next

End Sub

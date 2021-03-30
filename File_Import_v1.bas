Attribute VB_Name = "Module1"
Option Explicit

Sub WeldPlanMerger_v13()
Form.Show


End Sub

Sub lstToCollection(lst As MSForms.ListBox, col As Collection)
'// List Box�� ����� Collection ��ü�� ��ȯ
Dim i As Integer
    
    For i = 0 To lst.ListCount - 1
        col.Add lst.List(i, 0)
    
    Next
    


End Sub
Sub StatusLabel(lbl As MSForms.Label, strText As String)

lbl.Caption = strText

End Sub


Sub Listmove(lstOriginal As MSForms.ListBox, lstMove As MSForms.ListBox, Optional All As Boolean = False) '// list�� �׸� �̵��ϴ� Procedure ����
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
' ����ٷ� ���е� Line No.���� ���� ������ �и�
 Dim f As Variant
  f = Split(str, sepChar)
  If n > 0 And n - 1 <= UBound(f) Then
    splice = f(n - 1)
  Else
    splice = ""
  End If
  
End Function

Function TextCount(Text, Search As String) As Integer
   'Text�� ����ִ� Search�� ����
   Dim i As Integer
   TextCount = 0
   For i = 1 To Len(Text)
      If Mid(Text, i, 1) = Search Then TextCount = TextCount + 1
   Next
End Function

Sub WeldPlanMerge(wbLists As Collection, Optional optSet As String = "Default")
'// �� Workbook�� ���鼭 Range���� ����/�ٿ��ֱ� �ݺ�
Dim i As Integer '// Progress Bar Width ������ ����

Dim no As Integer '// list ���� Count
no = wbLists.count
Dim NoRow As Integer
Dim strFileName As String
Dim strDwg As String
Dim noSep As Integer '// "\" ���� count
Dim wbList As Variant
Dim wb As Workbook
Dim wbTemp As Workbook
Dim startTime As Single
startTime = Timer
Dim finishTime As Single

'// Set wbTemp = Workbooks.Open("****.xlsx", False, False)
'//range ��ü ����

'// �� WeldPlan ���� ������ Column�� Range ��ü�� ����



Dim rng As Range '// Range Copy
Dim c As Range '// �ݺ��� ���� ���� Range ��ü ����

'// ���� ������ ���� ����

Dim endRow As Integer

Dim NoRows As Integer



'// Directory ����� ����

ChDir ThisWorkbook.Path & "\" '//���� Directory�� ����

On Error Resume Next
Application.ScreenUpdating = False
Application.StatusBar = "�غ�"

If wbLists.count = 0 Then
    Exit Sub
    
Else
    If optSet = "Default" Then '// Setting�� "Default" �� ���
        i = 1
        For Each wbList In wbLists
        '// WorkBook ��ü Wb�� �Ҵ�
        Set wb = Workbooks.Open(wbList, False, True)
        noSep = TextCount(wbList, "\") '// directory Seperater ���� Count
        strFileName = splice(wbList, noSep + 1, "\")
        Application.StatusBar = strFileName & " - File ���� " & "(" & i & "of" & no & ")"
        On Error Resume Next
        Application.ScreenUpdating = False
            NoRow = Range("A2").End(xlDown).Row
            NoRows = NoRow - 2 + 1
            '// Workbook���� �� Column ������ �Ҵ�
            Application.StatusBar = strFileName & " - Data �б� " & "(" & i & "of" & no & ")"
           
            '//Template ���� Ȱ��ȭ
            wbTemp.Activate
            '// ���� ����
            Set rng = wbTemp.Sheets(1).Range(Range("H1048576").End(xlUp).Offset(1, 0), Range("H1048576").End(xlUp).Offset(NoRows, 0))
            Application.StatusBar = strFileName & " - Data �����ϱ� " & "(" & i & "of" & no & ")"
            '// JOINT ����
            rngJoint.Copy rng
            '// RAW_MATERIAL ����
            rngMatl.Copy rng.Offset(0, 1)
            '// SIZE ����
            rngSize.Copy rng.Offset(0, 2)
            '// SCHEDULE ����
            rngSch.Copy rng.Offset(0, 3)
            '// WELD TYPE ����
            rngType.Copy rng.Offset(0, 4)
            '// SHOP_FIELD ����
            rngSF.Copy rng.Offset(0, 5)
            '// SPOOL_NO ����
            rngSpool.Copy rng.Offset(0, 6)
            '// CON_SPOOL ����
            rngConSpool.Copy rng.Offset(0, 7)
            '// CON_ISONO ����
            rngConIso.Copy rng.Offset(0, 8)
            '// SPEC ����
            rngSpec.Copy rng.Offset(0, -1)
            '// CON_ISONO�ݺ��� ����
            For Each c In rng.Offset(0, 8)
            c.Value = splice(c, 4, "__")
            Next
             Application.StatusBar = strFileName & " - FMCS�� Data �Է� " & "(" & i & "of" & no & ")"
            '// Sbnm �Է�
            rng.Offset(0, -7).Value = "AG"
        
            '// Area �Է�
            rng.Offset(0, -6).Value = splice(strFileName, 2, "__")
            
            '// Book �Է�
            rng.Offset(0, -5).Value = splice(strFileName, 3, "__")
            
            '// Dwg �Է�
            strDwg = splice(strFileName, 4, "__") '// ���ϸ��� Splicing
            strDwg = splice(strDwg, 6, "-") '// ���ϸ��� Splicing
            strDwg = splice(strDwg, 1, "__") '// ���ϸ��� Splicing
            rng.Offset(0, -4).Value = strDwg
            '// Fluid �Է�
            rng.Offset(0, -3).Value = splice(strFileName, 3, "__")
            
            '// Line�Է�
            rng.Offset(0, -2).Value = splice(strFileName, 3, "__") & "-" & strDwg
            
            Application.StatusBar = strFileName & " - �Է� �Ϸ� " & "(" & i & "of" & no & ")"
            
            wb.Close False
            i = i + 1
            Application.Wait (Now + TimeValue("0:00:01"))
            Next
        End If
  
    
    
End If

'//�۾� �Ϸ� �� Message ��� �� ����
Application.StatusBar = False
finishTime = Timer
MsgBox "�� " & no & "�� Merge �Ϸ�" & Chr(13) & Chr(10) & "�ҿ�ð� : " & finishTime - startTime & "sec"
wbTemp.SaveAs "C:\WeldPlanMergeExport" & "\" & strSaveAs
End Sub



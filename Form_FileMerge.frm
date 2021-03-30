VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Weld Plan Merge v14"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14715
   OleObjectBlob   =   "Form_FileMerge.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbClear_Click()
lstWpUnSelect.Clear

End Sub

Private Sub cmbExportClear_Click()
lstWpSelect.Clear
End Sub

Private Sub cmbFiles_Click()
'//파일 추가 버튼 procedure
Set lstImport = Nothing '// Import collection 개체 삭제

Dim strFileName As String

Dim noSep As Integer

Dim WeldLists As Variant, Files As Variant

ChDir ThisWorkbook.Path & "\"

Dim noWpFiles As Integer

WeldLists = Application.GetOpenFilename("Weld Plan Report, *.xlsx", Title:="Weld Plan 선택", MultiSelect:=True)

On Error Resume Next
'// 화면 Update False 설정

Application.ScreenUpdating = False
'// Caption 입력

lblStatus2.Caption = "목록 읽는 중"
'// 파일 추가 하지 않았을 경우 Prcedure 종료

If WeldLists <> False Then Else Exit Sub

For Each Files In WeldLists
     If InStr(Files, "__WP__") Then '//File명 검사
        noWpFiles = noWpFiles + 1 '//Weld Plan File 수량 하나씩 추가
        noSep = TextCount(Files, "\") '// directory Seperater 개수 Cout
        strFileName = splice(Files, noSep + 1, "W") '// 파일명만 분리
        Debug.Print strFileName
        lblStatus2.Caption = "Import List에 파일 추가 중 : " & strFileName '// 추가중인 파일명 표시
        lstImport.Add (Files) '// collection 개체에 추가
     End If
     
    
Next

lblStatus2.Caption = "File 추가 완료 : 총 " & noWpFiles & "개"

End Sub

Private Sub cmbImport_Click()
'// import collection List에 표시
Dim file As Variant

If lstImport.count = 0 Then
    Exit Sub '// 수량이 0이면 Sub 종료
Else
    For Each file In lstImport
        lstWpUnSelect.AddItem (file)
    
    Next
End If
End Sub

Private Sub cmbImportClear_Click()
lstWpUnSelect.Clear
End Sub

Private Sub cmbMergeSave_Click()

'// Listbox의 Data collection으로 변경
Call lstToCollection(lstWpSelect, lstExport)

If lstExport.count = 0 And strSaveAs = "" Then

    MsgBox "File이 없습니다." & Chr(13) & Chr(10) & "File Name 미지정", vbOKOnly, "Error"
    
    Exit Sub


ElseIf lstExport.count <> 0 And strSaveAs = "" Then
    MsgBox "File Name 미지정", vbOKOnly, "File Name Error"
    Exit Sub

ElseIf lstExport.count = 0 And strSaveAs <> "" Then
    MsgBox "File이 없습니다.", vbOKOnly, "File Name Error"
    Exit Sub
Else
    

    Call WeldPlanMerge(lstExport)
    Unload Me
End If

End Sub

Private Sub cmbProject_Change()
    lblProject2.Caption = Me.cmbProject.Value
End Sub

Private Sub cmbRefresh_Click()
'// 파일이 추가 되었을 경우 lblTempFolderCheck.Caption 수정
    strTemp = "C:\WeldPlanMergeTemplate"
    strExport = "C:\WeldPlanMergeExport"
    strFname = "WeldPlanExcelUploadTemplate.xlsx"
    If Len(Dir(strTemp + "\" + strFname, vbDirectory)) = 0 Then
        lblTempFileCheck.Caption = "Template 파일이 없습니다"
    Else
        lblTempFileCheck.Caption = "Template 파일 확인 완료"
        
    End If
    If Len(Dir(strTemp, vbDirectory)) = 0 And Len(Dir(strExport, vbDirectory)) = 0 Then
        lblTempFolderCheck.Caption = "Template/Export 폴더 없음"
    ElseIf Len(Dir(strTemp, vbDirectory)) = 0 And Len(Dir(strExport, vbDirectory)) <> 0 Then
        lblTempFolderCheck.Caption = "Template 폴더 없음"
    ElseIf Len(Dir(strTemp, vbDirectory)) <> 0 And Len(Dir(strExport, vbDirectory)) = 0 Then
        lblTempFolderCheck.Caption = "Export 폴더 없음"
    
    Else
    
        lblTempFolderCheck.Caption = "Folder 확인 완료"
    
    End If

End Sub

Private Sub cmbSaveAs_Click()
If optSave = True Then
    strSaveAs = Format(DateTime.Now, "yymmddhhmmss") & "_" & "WeldPlanMerge.xlsx"
    lblSaveAs = strSaveAs
ElseIf optSaveAs = True And txtFileName.Value <> "" Then
    strSaveAs = txtFileName.Value & ".xlsx"
    lblSaveAs = strSaveAs
Else
    strSaveAs = Format(DateTime.Now, "yymmddhhmmss") & "_" & "WeldPlanMerge.xlsx"
    lblSaveAs = strSaveAs

End If
End Sub

Private Sub cmbSelectAll_Click()
Call Listmove(lstWpUnSelect, lstWpSelect, True)
End Sub

Private Sub cmbUnSelectAll_Click()
Call Listmove(lstWpSelect, lstWpUnSelect, True)
End Sub

Private Sub cmbUnSelectOne_Click()
Call Listmove(lstWpSelect, lstWpUnSelect, False)
End Sub

Private Sub cmdSelectOne_Click()

Call Listmove(lstWpUnSelect, lstWpSelect, False)

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub Frame4_Click()

End Sub

Private Sub lblProgressBar_Click()

End Sub

Private Sub lblSaveAs_Click()

End Sub

Private Sub lblTemplateDirectory_Click()

End Sub

Private Sub lstWpSelect_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub txtFileName_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
'// Form 초기화 추가
    
    Set lstImport = Nothing '//전역변수 초기화
    Set lstExport = Nothing '//전역변수 초기화
    
    Me.lstWpUnSelect.Clear
    Me.lstWpSelect.Clear
    Me.cmbProject.AddItem "Default"
    Me.cmbProject.AddItem "SE6666"
    Me.cmbProject.Value = "Default"
    Me.optSave = True
    strSaveAs = "" '//전역변수 초기화
    Me.lblSaveAs = strSaveAs '//전역변수 초기 설정
    strTemp = "C:\WeldPlanMergeTemplate" '//전역변수 초기 설정
    strExport = "C:\WeldPlanMergeExport" '//전역변수 초기 설정
    strFname = "WeldPlanExcelUploadTemplate.xlsx" '//전역변수 초기 설정
    '//경로 확인
    
    Debug.Print Len(Dir(strTemp, vbDirectory))
    If Len(Dir(strTemp, vbDirectory)) = 0 And Len(Dir(strExport, vbDirectory)) = 0 Then
        MkDir (strTemp)
        MkDir (strExport)
        lblTempFolderCheck.Caption = "Template/Export 폴더 생성 완료"
    ElseIf Len(Dir(strTemp, vbDirectory)) <> 0 And Len(Dir(strExport, vbDirectory)) = 0 Then
        MkDir (strExport)
        lblTempFolderCheck.Caption = "Export 폴더 생성 완료"
    
    ElseIf Len(Dir(strTemp, vbDirectory)) = 0 And Len(Dir(strExport, vbDirectory)) <> 0 Then
        MkDir (strTemp)
        lblTempFolderCheck.Caption = "Template 폴더 생성 완료"
    
    Else
    
        lblTempFolderCheck.Caption = "Folder 확인 완료"
    
    End If
    
    '// 파일 확인
    If Len(Dir(strTemp + "\" + strFname, vbDirectory)) = 0 Then
        lblTempFileCheck.Caption = "Template 파일이 없습니다"
    Else
        lblTempFileCheck.Caption = "Template 파일 확인 완료"
        
    End If
'// Listbox scrollbar 활성화
    With Me.lstWpUnSelect
        .ColumnWidths = 1000
    End With
'// Listbox scrollbar 활성화
    With Me.lstWpSelect
        .ColumnWidths = 1000
    End With

End Sub

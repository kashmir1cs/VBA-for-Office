VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Weld Plan Merge v14"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14715
   OleObjectBlob   =   "Form_FileMerge.frx":0000
   StartUpPosition =   1  '������ ���
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
'//���� �߰� ��ư procedure
Set lstImport = Nothing '// Import collection ��ü ����

Dim strFileName As String

Dim noSep As Integer

Dim WeldLists As Variant, Files As Variant

ChDir ThisWorkbook.Path & "\"

Dim noWpFiles As Integer

WeldLists = Application.GetOpenFilename("Weld Plan Report, *.xlsx", Title:="Weld Plan ����", MultiSelect:=True)

On Error Resume Next
'// ȭ�� Update False ����

Application.ScreenUpdating = False
'// Caption �Է�

lblStatus2.Caption = "��� �д� ��"
'// ���� �߰� ���� �ʾ��� ��� Prcedure ����

If WeldLists <> False Then Else Exit Sub

For Each Files In WeldLists
     If InStr(Files, "__WP__") Then '//File�� �˻�
        noWpFiles = noWpFiles + 1 '//Weld Plan File ���� �ϳ��� �߰�
        noSep = TextCount(Files, "\") '// directory Seperater ���� Cout
        strFileName = splice(Files, noSep + 1, "W") '// ���ϸ� �и�
        Debug.Print strFileName
        lblStatus2.Caption = "Import List�� ���� �߰� �� : " & strFileName '// �߰����� ���ϸ� ǥ��
        lstImport.Add (Files) '// collection ��ü�� �߰�
     End If
     
    
Next

lblStatus2.Caption = "File �߰� �Ϸ� : �� " & noWpFiles & "��"

End Sub

Private Sub cmbImport_Click()
'// import collection List�� ǥ��
Dim file As Variant

If lstImport.count = 0 Then
    Exit Sub '// ������ 0�̸� Sub ����
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

'// Listbox�� Data collection���� ����
Call lstToCollection(lstWpSelect, lstExport)

If lstExport.count = 0 And strSaveAs = "" Then

    MsgBox "File�� �����ϴ�." & Chr(13) & Chr(10) & "File Name ������", vbOKOnly, "Error"
    
    Exit Sub


ElseIf lstExport.count <> 0 And strSaveAs = "" Then
    MsgBox "File Name ������", vbOKOnly, "File Name Error"
    Exit Sub

ElseIf lstExport.count = 0 And strSaveAs <> "" Then
    MsgBox "File�� �����ϴ�.", vbOKOnly, "File Name Error"
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
'// ������ �߰� �Ǿ��� ��� lblTempFolderCheck.Caption ����
    strTemp = "C:\WeldPlanMergeTemplate"
    strExport = "C:\WeldPlanMergeExport"
    strFname = "WeldPlanExcelUploadTemplate.xlsx"
    If Len(Dir(strTemp + "\" + strFname, vbDirectory)) = 0 Then
        lblTempFileCheck.Caption = "Template ������ �����ϴ�"
    Else
        lblTempFileCheck.Caption = "Template ���� Ȯ�� �Ϸ�"
        
    End If
    If Len(Dir(strTemp, vbDirectory)) = 0 And Len(Dir(strExport, vbDirectory)) = 0 Then
        lblTempFolderCheck.Caption = "Template/Export ���� ����"
    ElseIf Len(Dir(strTemp, vbDirectory)) = 0 And Len(Dir(strExport, vbDirectory)) <> 0 Then
        lblTempFolderCheck.Caption = "Template ���� ����"
    ElseIf Len(Dir(strTemp, vbDirectory)) <> 0 And Len(Dir(strExport, vbDirectory)) = 0 Then
        lblTempFolderCheck.Caption = "Export ���� ����"
    
    Else
    
        lblTempFolderCheck.Caption = "Folder Ȯ�� �Ϸ�"
    
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
'// Form �ʱ�ȭ �߰�
    
    Set lstImport = Nothing '//�������� �ʱ�ȭ
    Set lstExport = Nothing '//�������� �ʱ�ȭ
    
    Me.lstWpUnSelect.Clear
    Me.lstWpSelect.Clear
    Me.cmbProject.AddItem "Default"
    Me.cmbProject.AddItem "SE6666"
    Me.cmbProject.Value = "Default"
    Me.optSave = True
    strSaveAs = "" '//�������� �ʱ�ȭ
    Me.lblSaveAs = strSaveAs '//�������� �ʱ� ����
    strTemp = "C:\WeldPlanMergeTemplate" '//�������� �ʱ� ����
    strExport = "C:\WeldPlanMergeExport" '//�������� �ʱ� ����
    strFname = "WeldPlanExcelUploadTemplate.xlsx" '//�������� �ʱ� ����
    '//��� Ȯ��
    
    Debug.Print Len(Dir(strTemp, vbDirectory))
    If Len(Dir(strTemp, vbDirectory)) = 0 And Len(Dir(strExport, vbDirectory)) = 0 Then
        MkDir (strTemp)
        MkDir (strExport)
        lblTempFolderCheck.Caption = "Template/Export ���� ���� �Ϸ�"
    ElseIf Len(Dir(strTemp, vbDirectory)) <> 0 And Len(Dir(strExport, vbDirectory)) = 0 Then
        MkDir (strExport)
        lblTempFolderCheck.Caption = "Export ���� ���� �Ϸ�"
    
    ElseIf Len(Dir(strTemp, vbDirectory)) = 0 And Len(Dir(strExport, vbDirectory)) <> 0 Then
        MkDir (strTemp)
        lblTempFolderCheck.Caption = "Template ���� ���� �Ϸ�"
    
    Else
    
        lblTempFolderCheck.Caption = "Folder Ȯ�� �Ϸ�"
    
    End If
    
    '// ���� Ȯ��
    If Len(Dir(strTemp + "\" + strFname, vbDirectory)) = 0 Then
        lblTempFileCheck.Caption = "Template ������ �����ϴ�"
    Else
        lblTempFileCheck.Caption = "Template ���� Ȯ�� �Ϸ�"
        
    End If
'// Listbox scrollbar Ȱ��ȭ
    With Me.lstWpUnSelect
        .ColumnWidths = 1000
    End With
'// Listbox scrollbar Ȱ��ȭ
    With Me.lstWpSelect
        .ColumnWidths = 1000
    End With

End Sub

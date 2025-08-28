VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_SheetRenamer 
   Caption         =   "�V�[�g���ꊇ�ύX�c�[��"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   OleObjectBlob   =   "UserForm_SheetRenamer.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm_SheetRenamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ���W���[�����x���ϐ��i�A�ԊǗ��p�j
Private sequenceCounter As Integer

Private Sub UserForm_Initialize()
    ' �t�H�[��������
    Me.caption = "�V�[�g���ꊇ�ύX�c�[��"
    
    ' �V�[�g�ꗗ��ǂݍ���
    Call LoadSheetList
    
    ' �f�t�H���g�ݒ�
    optSequence.value = True
    Call UpdateInterface
End Sub

Private Sub LoadSheetList()
    ' ���݂̃��[�N�u�b�N�̃V�[�g�ꗗ��ǂݍ���
    Dim ws As Worksheet
    lstSheets.Clear
    
    For Each ws In ActiveWorkbook.Worksheets
        lstSheets.AddItem ws.Name
    Next ws
    
    ' �S�I���`�F�b�N�{�b�N�X�̏������
    chkSelectAll.value = True
    Call chkSelectAll_Click
End Sub

Private Sub optIndividual_Click()
    Call UpdateInterface
End Sub

Private Sub optReplace_Click()
    Call UpdateInterface
End Sub

Private Sub optPrefix_Click()
    Call UpdateInterface
End Sub

Private Sub optSuffix_Click()
    Call UpdateInterface
End Sub

Private Sub optSequence_Click()
    Call UpdateInterface
End Sub

Private Sub UpdateInterface()
    ' �I�����ꂽ�ύX���@�ɉ����ăC���^�[�t�F�[�X���X�V
    
    ' �S�Ĕ�\���ɂ��Ă��烊�Z�b�g
    lblFind.Visible = False
    txtFind.Visible = False
    lblReplace.Visible = False
    txtReplace.Visible = False
    lblPrefix.Visible = False
    txtPrefix.Visible = False
    lblSuffix.Visible = False
    txtSuffix.Visible = False
    
    ' �A�Ԋ֘A�R���g���[���i���݂���ꍇ�̂݁j
    Call HideSequenceControls
    
    If optIndividual.value Then
        ' �ʎw�胂�[�h
        lblInstruction.caption = "�ύX�������V�[�g��I�����āA�v���r���[�ŐV�������O����͂��Ă�������"
        
    ElseIf optReplace.value Then
        ' �u�����[�h
        lblFind.Visible = True
        txtFind.Visible = True
        lblReplace.Visible = True
        txtReplace.Visible = True
        lblInstruction.caption = "�u��������������ƐV�������������͂��Ă�������"
        
    ElseIf optPrefix.value Then
        ' �v���t�B�b�N�X�ǉ����[�h
        lblPrefix.Visible = True
        txtPrefix.Visible = True
        lblInstruction.caption = "�V�[�g���̐擪�ɒǉ����镶�������͂��Ă�������"
        
    ElseIf optSuffix.value Then
        ' �T�t�B�b�N�X�ǉ����[�h
        lblSuffix.Visible = True
        txtSuffix.Visible = True
        lblInstruction.caption = "�V�[�g���̖����ɒǉ����镶�������͂��Ă�������"
        
    ElseIf ControlExists("optSequence") And optSequence.value Then
        ' �A�ԃ��[�h�i�R���g���[�������݂���ꍇ�̂݁j
        Call ShowSequenceControls
        lblInstruction.caption = "�A�Ԃ̎�ށA�ʒu�A�J�n�ԍ���ݒ肵�Ă�������"
        
        ' �A�ԃR���g���[���̏�����
        Call InitializeSequenceControls
    End If
End Sub

Private Function ControlExists(controlName As String) As Boolean
    ' �t�H�[����ɃR���g���[�������݂��邩�`�F�b�N
    On Error Resume Next
    Dim ctrl As Control
    Set ctrl = Me.Controls(controlName)
    ControlExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub HideSequenceControls()
    ' �A�Ԋ֘A�R���g���[�����\���i���݂���ꍇ�̂݁j
    If ControlExists("lblSequenceType") Then lblSequenceType.Visible = False
    If ControlExists("cmbSequenceType") Then cmbSequenceType.Visible = False
    If ControlExists("lblSequencePosition") Then lblSequencePosition.Visible = False
    If ControlExists("cmbSequencePosition") Then cmbSequencePosition.Visible = False
    If ControlExists("lblSequenceStart") Then lblSequenceStart.Visible = False
    If ControlExists("txtSequenceStart") Then txtSequenceStart.Visible = False
    If ControlExists("lblSequenceSpan") Then lblSequenceSpan.Visible = False
    If ControlExists("txtSequenceSpan") Then txtSequenceSpan.Visible = False
    If ControlExists("lblSequenceDigits") Then lblSequenceDigits.Visible = False
    If ControlExists("txtSequenceDigits") Then txtSequenceDigits.Visible = False
End Sub

Private Sub ShowSequenceControls()
    ' �A�Ԋ֘A�R���g���[����\���i���݂���ꍇ�̂݁j
    If ControlExists("lblSequenceType") Then lblSequenceType.Visible = True
    If ControlExists("cmbSequenceType") Then cmbSequenceType.Visible = True
    If ControlExists("lblSequencePosition") Then lblSequencePosition.Visible = True
    If ControlExists("cmbSequencePosition") Then cmbSequencePosition.Visible = True
    If ControlExists("lblSequenceStart") Then lblSequenceStart.Visible = True
    If ControlExists("txtSequenceStart") Then txtSequenceStart.Visible = True
    If ControlExists("lblSequenceSpan") Then lblSequenceSpan.Visible = True
    If ControlExists("txtSequenceSpan") Then txtSequenceSpan.Visible = True
    If ControlExists("lblSequenceDigits") Then lblSequenceDigits.Visible = True
    If ControlExists("txtSequenceDigits") Then txtSequenceDigits.Visible = True
End Sub

Private Sub chkSelectAll_Click()
    ' �S�I��/�S����
    Dim i As Integer
    For i = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(i) = chkSelectAll.value
    Next i
End Sub

Private Sub btnPreview_Click()
    ' �v���r���[�𐶐�
    Call GeneratePreview
End Sub

Private Sub GeneratePreview()
    ' �I�����ꂽ�ύX���@�Ɋ�Â��ăv���r���[�𐶐�
    Dim i As Integer
    Dim oldName As String
    Dim newName As String
    
    lstPreview.Clear
    
    ' �A�ԃ��[�h�̏ꍇ�A�J�E���^�[�����Z�b�g
    If ControlExists("optSequence") And optSequence.value Then
        Call ResetSequenceCounter
    End If
    
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then
            oldName = lstSheets.List(i)
            newName = GetNewSheetName(oldName)
            
            ' �v���r���[�\�� (���� �� �V��)
            lstPreview.AddItem oldName & " �� " & newName
        End If
    Next i
    
    ' �ʎw�胂�[�h�̏ꍇ�A�v���r���[���X�g��ҏW�\�ɂ���
    If optIndividual.value Then
        Call EnableIndividualEdit
    End If
End Sub

Private Sub ResetSequenceCounter()
    ' �A�ԃJ�E���^�[�����Z�b�g
    sequenceCounter = 1 ' ���W���[�����x���ϐ���1�Ƀ��Z�b�g
End Sub

Private Sub InitializeSequenceControls()
    ' �A�ԃR���g���[���̏������i���݂���ꍇ�̂݁j
    
    If ControlExists("cmbSequenceType") Then
        ' �A�ԃ^�C�v�̐ݒ�
        cmbSequenceType.Clear
        cmbSequenceType.AddItem "���� (1, 2, 3...)"
        cmbSequenceType.AddItem "�����[���p�f�B���O (01, 02, 03...)"
        cmbSequenceType.AddItem "�������A���t�@�x�b�g (a, b, c...)"
        cmbSequenceType.AddItem "�啶���A���t�@�x�b�g (A, B, C...)"
        cmbSequenceType.ListIndex = 1 ' �f�t�H���g�̓[���p�f�B���O
    End If
    
    If ControlExists("cmbSequencePosition") Then
        ' �A�Ԉʒu�̐ݒ�
        cmbSequencePosition.Clear
        cmbSequencePosition.AddItem "�v���t�B�b�N�X�i�擪�j"
        cmbSequencePosition.AddItem "�T�t�B�b�N�X�i�����j"
        cmbSequencePosition.AddItem "�u���i���S�u���j"
        cmbSequencePosition.ListIndex = 0 ' �f�t�H���g�̓v���t�B�b�N�X
    End If
    
    ' �f�t�H���g�l�̐ݒ�
    If ControlExists("txtSequenceStart") Then txtSequenceStart.Text = "1"
    If ControlExists("txtSequenceSpan") Then txtSequenceSpan.Text = "1"
    If ControlExists("txtSequenceDigits") Then txtSequenceDigits.Text = "2"
End Sub

Private Function GetNewSheetName(oldName As String) As String
    ' �I�����ꂽ���@�Ɋ�Â��ĐV�����V�[�g���𐶐�
    Dim newName As String
    
    If optIndividual.value Then
        ' �ʎw�� - �Ƃ肠�������̖��O��Ԃ��i��ŕҏW�\�j
        newName = oldName
        
    ElseIf optReplace.value Then
        ' �u��
        newName = Replace(oldName, txtFind.Text, txtReplace.Text)
        
    ElseIf optPrefix.value Then
        ' �v���t�B�b�N�X�ǉ�
        newName = txtPrefix.Text & oldName
        
    ElseIf optSuffix.value Then
        ' �T�t�B�b�N�X�ǉ�
        newName = oldName & txtSuffix.Text
        
    ElseIf ControlExists("optSequence") And optSequence.value Then
        ' �A�Ԓǉ��i�R���g���[�������݂���ꍇ�̂݁j
        newName = GetSequenceName(oldName)
    Else
        ' �f�t�H���g�i�ύX�Ȃ��j
        newName = oldName
    End If
    
    GetNewSheetName = newName
End Function

Private Function GetSequenceName(oldName As String) As String
    ' �A�Ԃ𐶐����ăV�[�g�����쐬
    Dim sequenceText As String
    Dim newName As String
    Dim startNum As Integer
    Dim spanNum As Integer
    Dim digitCount As Integer
    Dim currentValue As Integer
    
    ' �J�n�ԍ��ƃX�p���̎擾
    startNum = 1
    spanNum = 1
    
    If ControlExists("txtSequenceStart") Then
        If IsNumeric(txtSequenceStart.Text) Then
            startNum = Val(txtSequenceStart.Text)
        End If
    End If
    
    If ControlExists("txtSequenceSpan") Then
        If IsNumeric(txtSequenceSpan.Text) Then
            spanNum = Val(txtSequenceSpan.Text)
        End If
    End If
    
    ' ���ݒl���v�Z�F�J�n�ԍ� + (�J�E���^�[ - 1) �~ �X�p��
    currentValue = startNum + (sequenceCounter - 1) * spanNum
    
    ' �A�ԕ�����𐶐�
    Select Case cmbSequenceType.ListIndex
        Case 0 ' ���� (1, 2, 3...)
            sequenceText = CStr(currentValue)
            
        Case 1 ' �����[���p�f�B���O (01, 02, 03...)
            digitCount = 2 ' �f�t�H���g�l
            If ControlExists("txtSequenceDigits") Then
                If IsNumeric(txtSequenceDigits.Text) Then
                    digitCount = Val(txtSequenceDigits.Text)
                End If
            End If
            If digitCount < 1 Then digitCount = 2
            sequenceText = Format(currentValue, String(digitCount, "0"))
            
        Case 2 ' �������A���t�@�x�b�g (a, b, c...)
            sequenceText = GetExcelColumnName(currentValue, False)
            
        Case 3 ' �啶���A���t�@�x�b�g (A, B, C...)
            sequenceText = GetExcelColumnName(currentValue, True)
            
        Case Else
            sequenceText = CStr(currentValue)
    End Select
    
    ' �ʒu�ɉ����ăV�[�g���𐶐�
    Select Case cmbSequencePosition.ListIndex
        Case 0 ' �v���t�B�b�N�X�i�擪�j
            newName = sequenceText & oldName
            
        Case 1 ' �T�t�B�b�N�X�i�����j
            newName = oldName & sequenceText
            
        Case 2 ' �u���i���S�u���j
            newName = sequenceText
            
        Case Else
            newName = sequenceText & oldName
    End Select
    
    ' �J�E���^�[�𑝉�
    sequenceCounter = sequenceCounter + 1
    
    GetSequenceName = newName
End Function

Private Function GetExcelColumnName(columnNumber As Integer, isUpperCase As Boolean) As String
    ' Excel�̗񖼂Ɠ��l�̃A���t�@�x�b�g�\���𐶐�
    ' 1:A, 2:B, ..., 26:Z, 27:AA, 28:AB, ...
    Dim result As String
    Dim tempNum As Integer
    Dim baseChar As String
    
    If isUpperCase Then
        baseChar = "A"
    Else
        baseChar = "a"
    End If
    
    tempNum = columnNumber
    
    Do
        tempNum = tempNum - 1 ' 0�x�[�X�ɕϊ�
        result = Chr(Asc(baseChar) + (tempNum Mod 26)) & result
        tempNum = tempNum \ 26
    Loop While tempNum > 0
    
    GetExcelColumnName = result
End Function

Private Sub EnableIndividualEdit()
    ' �ʎw�胂�[�h�Ńv���r���[���X�g��ҏW�\�ɂ���
    MsgBox "�ʎw�胂�[�h�ł́A���s���Ɋe�V�[�g�̐V�������O���ʂɓ��͂ł��܂��B", vbInformation
End Sub

Private Sub btnExecute_Click()
    ' �ύX�����s
    If lstPreview.ListCount = 0 Then
        MsgBox "�܂��v���r���[�𐶐����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' �m�F���b�Z�[�W
    Dim result As VbMsgBoxResult
    result = MsgBox("�I�����ꂽ�V�[�g����ύX���܂����H" & vbCrLf & _
                    "���̑���͌��ɖ߂��܂���B", vbYesNo + vbQuestion)
    
    If result = vbYes Then
        Call ExecuteRename
    End If
End Sub

Private Sub ExecuteRename()
    ' ���ۂɃV�[�g����ύX
    Dim i As Integer
    Dim oldName As String
    Dim newName As String
    Dim ws As Worksheet
    Dim errorCount As Integer
    Dim successCount As Integer
    Dim skipFlag As Boolean
    
    errorCount = 0
    successCount = 0
    
    Application.ScreenUpdating = False
    
    ' �A�ԃ��[�h�̏ꍇ�A�J�E���^�[�����Z�b�g
    If ControlExists("optSequence") And optSequence.value Then
        Call ResetSequenceCounter
    End If
    
    For i = 0 To lstSheets.ListCount - 1
        skipFlag = False
        
        If lstSheets.Selected(i) Then
            oldName = lstSheets.List(i)
            
            ' �ʎw�胂�[�h�̏ꍇ�A�V�������O����͂��Ă��炤
            If optIndividual.value Then
                newName = InputBox("�V�����V�[�g������͂��Ă�������:", "�V�[�g���ύX", oldName)
                If newName = "" Or newName = oldName Then
                    skipFlag = True ' �X�L�b�v
                End If
            Else
                newName = GetNewSheetName(oldName)
            End If
            
            ' �X�L�b�v�t���O���`�F�b�N
            If Not skipFlag Then
                ' �V�[�g���̗L�����`�F�b�N
                If Not IsValidSheetName(newName) Then
                    MsgBox "�����ȃV�[�g���ł�: " & newName & vbCrLf & _
                           "�ȉ��̕����͎g�p�ł��܂���: \ / ? * [ ] :", vbExclamation
                    errorCount = errorCount + 1
                    skipFlag = True
                End If
            End If
            
            If Not skipFlag Then
                ' �d���`�F�b�N
                If sheetExists(newName) And newName <> oldName Then
                    MsgBox "�����̃V�[�g�����ɑ��݂��܂�: " & newName, vbExclamation
                    errorCount = errorCount + 1
                    skipFlag = True
                End If
            End If
            
            If Not skipFlag Then
                ' �V�[�g���ύX
                On Error Resume Next
                Set ws = ActiveWorkbook.Worksheets(oldName)
                ws.Name = newName
                
                If Err.Number = 0 Then
                    successCount = successCount + 1
                Else
                    errorCount = errorCount + 1
                    MsgBox "�V�[�g���̕ύX�Ɏ��s���܂���: " & oldName & " �� " & newName & vbCrLf & _
                           "�G���[: " & Err.description, vbExclamation
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    ' ���ʕ�
    MsgBox "��������" & vbCrLf & _
           "����: " & successCount & " ��" & vbCrLf & _
           "���s: " & errorCount & " ��", vbInformation
    
    ' �V�[�g�ꗗ���X�V
    Call LoadSheetList
    lstPreview.Clear
End Sub

Private Function IsValidSheetName(sheetName As String) As Boolean
    ' �V�[�g���̗L�������`�F�b�N
    Dim invalidChars As String
    Dim i As Integer
    
    invalidChars = "\/?*[]:"
    
    ' �󕶎��`�F�b�N
    If Trim(sheetName) = "" Then
        IsValidSheetName = False
        Exit Function
    End If
    
    ' �����`�F�b�N�iExcel��31�����܂Łj
    If Len(sheetName) > 31 Then
        IsValidSheetName = False
        Exit Function
    End If
    
    ' ���������`�F�b�N
    For i = 1 To Len(invalidChars)
        If InStr(sheetName, Mid(invalidChars, i, 1)) > 0 Then
            IsValidSheetName = False
            Exit Function
        End If
    Next i
    
    IsValidSheetName = True
End Function

Private Function sheetExists(sheetName As String) As Boolean
    ' �w�肳�ꂽ�V�[�g�����݂��邩�`�F�b�N
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    sheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub btnCancel_Click()
    ' �t�H�[�������
    Unload Me
End Sub

' �e�L�X�g�{�b�N�X�E�R���{�{�b�N�X�̕ύX�C�x���g
Private Sub txtFind_Change()
    If optReplace.value And txtFind.Text <> "" And txtReplace.Text <> "" Then
        Call GeneratePreview
    End If
End Sub

Private Sub txtReplace_Change()
    If optReplace.value And txtFind.Text <> "" And txtReplace.Text <> "" Then
        Call GeneratePreview
    End If
End Sub

Private Sub txtPrefix_Change()
    If optPrefix.value And txtPrefix.Text <> "" Then
        Call GeneratePreview
    End If
End Sub

Private Sub txtSuffix_Change()
    If optSuffix.value And txtSuffix.Text <> "" Then
        Call GeneratePreview
    End If
End Sub

' �A�Ԋ֘A�̕ύX�C�x���g
Private Sub cmbSequenceType_Change()
    If ControlExists("optSequence") And optSequence.value Then
        Call GeneratePreview
    End If
End Sub

Private Sub cmbSequencePosition_Change()
    If ControlExists("optSequence") And optSequence.value Then
        Call GeneratePreview
    End If
End Sub

Private Sub txtSequenceStart_Change()
    If ControlExists("optSequence") And optSequence.value And IsNumeric(txtSequenceStart.Text) Then
        Call GeneratePreview
    End If
End Sub

Private Sub txtSequenceSpan_Change()
    If ControlExists("optSequence") And optSequence.value And IsNumeric(txtSequenceSpan.Text) Then
        Call GeneratePreview
    End If
End Sub

Private Sub txtSequenceDigits_Change()
    If ControlExists("optSequence") And optSequence.value And IsNumeric(txtSequenceDigits.Text) Then
        Call GeneratePreview
    End If
End Sub

' �I�����ꂽ�V�[�g�̃o�b�N�A�b�v���쐬
Sub CreateSheetBackup()
    Dim ws As Worksheet
    Dim backupWb As Workbook
    Dim originalName As String
    
    Set backupWb = Workbooks.Add
    originalName = ActiveWorkbook.Name
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Copy After:=backupWb.Sheets(backupWb.Sheets.count)
    Next ws
    
    ' �ŏ��̋�V�[�g���폜
    Application.DisplayAlerts = False
    backupWb.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    backupWb.SaveAs ThisWorkbook.Path & "\�V�[�g���ύX�O�o�b�N�A�b�v_" & _
                    Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    
    MsgBox "�o�b�N�A�b�v���쐬���܂���: " & backupWb.Name, vbInformation
End Sub


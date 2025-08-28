VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_SheetRenamer 
   Caption         =   "シート名一括変更ツール"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   OleObjectBlob   =   "UserForm_SheetRenamer.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm_SheetRenamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' モジュールレベル変数（連番管理用）
Private sequenceCounter As Integer

Private Sub UserForm_Initialize()
    ' フォーム初期化
    Me.caption = "シート名一括変更ツール"
    
    ' シート一覧を読み込み
    Call LoadSheetList
    
    ' デフォルト設定
    optSequence.value = True
    Call UpdateInterface
End Sub

Private Sub LoadSheetList()
    ' 現在のワークブックのシート一覧を読み込み
    Dim ws As Worksheet
    lstSheets.Clear
    
    For Each ws In ActiveWorkbook.Worksheets
        lstSheets.AddItem ws.Name
    Next ws
    
    ' 全選択チェックボックスの初期状態
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
    ' 選択された変更方法に応じてインターフェースを更新
    
    ' 全て非表示にしてからリセット
    lblFind.Visible = False
    txtFind.Visible = False
    lblReplace.Visible = False
    txtReplace.Visible = False
    lblPrefix.Visible = False
    txtPrefix.Visible = False
    lblSuffix.Visible = False
    txtSuffix.Visible = False
    
    ' 連番関連コントロール（存在する場合のみ）
    Call HideSequenceControls
    
    If optIndividual.value Then
        ' 個別指定モード
        lblInstruction.caption = "変更したいシートを選択して、プレビューで新しい名前を入力してください"
        
    ElseIf optReplace.value Then
        ' 置換モード
        lblFind.Visible = True
        txtFind.Visible = True
        lblReplace.Visible = True
        txtReplace.Visible = True
        lblInstruction.caption = "置換したい文字列と新しい文字列を入力してください"
        
    ElseIf optPrefix.value Then
        ' プレフィックス追加モード
        lblPrefix.Visible = True
        txtPrefix.Visible = True
        lblInstruction.caption = "シート名の先頭に追加する文字列を入力してください"
        
    ElseIf optSuffix.value Then
        ' サフィックス追加モード
        lblSuffix.Visible = True
        txtSuffix.Visible = True
        lblInstruction.caption = "シート名の末尾に追加する文字列を入力してください"
        
    ElseIf ControlExists("optSequence") And optSequence.value Then
        ' 連番モード（コントロールが存在する場合のみ）
        Call ShowSequenceControls
        lblInstruction.caption = "連番の種類、位置、開始番号を設定してください"
        
        ' 連番コントロールの初期化
        Call InitializeSequenceControls
    End If
End Sub

Private Function ControlExists(controlName As String) As Boolean
    ' フォーム上にコントロールが存在するかチェック
    On Error Resume Next
    Dim ctrl As Control
    Set ctrl = Me.Controls(controlName)
    ControlExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub HideSequenceControls()
    ' 連番関連コントロールを非表示（存在する場合のみ）
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
    ' 連番関連コントロールを表示（存在する場合のみ）
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
    ' 全選択/全解除
    Dim i As Integer
    For i = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(i) = chkSelectAll.value
    Next i
End Sub

Private Sub btnPreview_Click()
    ' プレビューを生成
    Call GeneratePreview
End Sub

Private Sub GeneratePreview()
    ' 選択された変更方法に基づいてプレビューを生成
    Dim i As Integer
    Dim oldName As String
    Dim newName As String
    
    lstPreview.Clear
    
    ' 連番モードの場合、カウンターをリセット
    If ControlExists("optSequence") And optSequence.value Then
        Call ResetSequenceCounter
    End If
    
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then
            oldName = lstSheets.List(i)
            newName = GetNewSheetName(oldName)
            
            ' プレビュー表示 (旧名 → 新名)
            lstPreview.AddItem oldName & " → " & newName
        End If
    Next i
    
    ' 個別指定モードの場合、プレビューリストを編集可能にする
    If optIndividual.value Then
        Call EnableIndividualEdit
    End If
End Sub

Private Sub ResetSequenceCounter()
    ' 連番カウンターをリセット
    sequenceCounter = 1 ' モジュールレベル変数を1にリセット
End Sub

Private Sub InitializeSequenceControls()
    ' 連番コントロールの初期化（存在する場合のみ）
    
    If ControlExists("cmbSequenceType") Then
        ' 連番タイプの設定
        cmbSequenceType.Clear
        cmbSequenceType.AddItem "数字 (1, 2, 3...)"
        cmbSequenceType.AddItem "数字ゼロパディング (01, 02, 03...)"
        cmbSequenceType.AddItem "小文字アルファベット (a, b, c...)"
        cmbSequenceType.AddItem "大文字アルファベット (A, B, C...)"
        cmbSequenceType.ListIndex = 1 ' デフォルトはゼロパディング
    End If
    
    If ControlExists("cmbSequencePosition") Then
        ' 連番位置の設定
        cmbSequencePosition.Clear
        cmbSequencePosition.AddItem "プレフィックス（先頭）"
        cmbSequencePosition.AddItem "サフィックス（末尾）"
        cmbSequencePosition.AddItem "置換（完全置換）"
        cmbSequencePosition.ListIndex = 0 ' デフォルトはプレフィックス
    End If
    
    ' デフォルト値の設定
    If ControlExists("txtSequenceStart") Then txtSequenceStart.Text = "1"
    If ControlExists("txtSequenceSpan") Then txtSequenceSpan.Text = "1"
    If ControlExists("txtSequenceDigits") Then txtSequenceDigits.Text = "2"
End Sub

Private Function GetNewSheetName(oldName As String) As String
    ' 選択された方法に基づいて新しいシート名を生成
    Dim newName As String
    
    If optIndividual.value Then
        ' 個別指定 - とりあえず元の名前を返す（後で編集可能）
        newName = oldName
        
    ElseIf optReplace.value Then
        ' 置換
        newName = Replace(oldName, txtFind.Text, txtReplace.Text)
        
    ElseIf optPrefix.value Then
        ' プレフィックス追加
        newName = txtPrefix.Text & oldName
        
    ElseIf optSuffix.value Then
        ' サフィックス追加
        newName = oldName & txtSuffix.Text
        
    ElseIf ControlExists("optSequence") And optSequence.value Then
        ' 連番追加（コントロールが存在する場合のみ）
        newName = GetSequenceName(oldName)
    Else
        ' デフォルト（変更なし）
        newName = oldName
    End If
    
    GetNewSheetName = newName
End Function

Private Function GetSequenceName(oldName As String) As String
    ' 連番を生成してシート名を作成
    Dim sequenceText As String
    Dim newName As String
    Dim startNum As Integer
    Dim spanNum As Integer
    Dim digitCount As Integer
    Dim currentValue As Integer
    
    ' 開始番号とスパンの取得
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
    
    ' 現在値を計算：開始番号 + (カウンター - 1) × スパン
    currentValue = startNum + (sequenceCounter - 1) * spanNum
    
    ' 連番文字列を生成
    Select Case cmbSequenceType.ListIndex
        Case 0 ' 数字 (1, 2, 3...)
            sequenceText = CStr(currentValue)
            
        Case 1 ' 数字ゼロパディング (01, 02, 03...)
            digitCount = 2 ' デフォルト値
            If ControlExists("txtSequenceDigits") Then
                If IsNumeric(txtSequenceDigits.Text) Then
                    digitCount = Val(txtSequenceDigits.Text)
                End If
            End If
            If digitCount < 1 Then digitCount = 2
            sequenceText = Format(currentValue, String(digitCount, "0"))
            
        Case 2 ' 小文字アルファベット (a, b, c...)
            sequenceText = GetExcelColumnName(currentValue, False)
            
        Case 3 ' 大文字アルファベット (A, B, C...)
            sequenceText = GetExcelColumnName(currentValue, True)
            
        Case Else
            sequenceText = CStr(currentValue)
    End Select
    
    ' 位置に応じてシート名を生成
    Select Case cmbSequencePosition.ListIndex
        Case 0 ' プレフィックス（先頭）
            newName = sequenceText & oldName
            
        Case 1 ' サフィックス（末尾）
            newName = oldName & sequenceText
            
        Case 2 ' 置換（完全置換）
            newName = sequenceText
            
        Case Else
            newName = sequenceText & oldName
    End Select
    
    ' カウンターを増加
    sequenceCounter = sequenceCounter + 1
    
    GetSequenceName = newName
End Function

Private Function GetExcelColumnName(columnNumber As Integer, isUpperCase As Boolean) As String
    ' Excelの列名と同様のアルファベット表現を生成
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
        tempNum = tempNum - 1 ' 0ベースに変換
        result = Chr(Asc(baseChar) + (tempNum Mod 26)) & result
        tempNum = tempNum \ 26
    Loop While tempNum > 0
    
    GetExcelColumnName = result
End Function

Private Sub EnableIndividualEdit()
    ' 個別指定モードでプレビューリストを編集可能にする
    MsgBox "個別指定モードでは、実行時に各シートの新しい名前を個別に入力できます。", vbInformation
End Sub

Private Sub btnExecute_Click()
    ' 変更を実行
    If lstPreview.ListCount = 0 Then
        MsgBox "まずプレビューを生成してください。", vbExclamation
        Exit Sub
    End If
    
    ' 確認メッセージ
    Dim result As VbMsgBoxResult
    result = MsgBox("選択されたシート名を変更しますか？" & vbCrLf & _
                    "この操作は元に戻せません。", vbYesNo + vbQuestion)
    
    If result = vbYes Then
        Call ExecuteRename
    End If
End Sub

Private Sub ExecuteRename()
    ' 実際にシート名を変更
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
    
    ' 連番モードの場合、カウンターをリセット
    If ControlExists("optSequence") And optSequence.value Then
        Call ResetSequenceCounter
    End If
    
    For i = 0 To lstSheets.ListCount - 1
        skipFlag = False
        
        If lstSheets.Selected(i) Then
            oldName = lstSheets.List(i)
            
            ' 個別指定モードの場合、新しい名前を入力してもらう
            If optIndividual.value Then
                newName = InputBox("新しいシート名を入力してください:", "シート名変更", oldName)
                If newName = "" Or newName = oldName Then
                    skipFlag = True ' スキップ
                End If
            Else
                newName = GetNewSheetName(oldName)
            End If
            
            ' スキップフラグをチェック
            If Not skipFlag Then
                ' シート名の有効性チェック
                If Not IsValidSheetName(newName) Then
                    MsgBox "無効なシート名です: " & newName & vbCrLf & _
                           "以下の文字は使用できません: \ / ? * [ ] :", vbExclamation
                    errorCount = errorCount + 1
                    skipFlag = True
                End If
            End If
            
            If Not skipFlag Then
                ' 重複チェック
                If sheetExists(newName) And newName <> oldName Then
                    MsgBox "同名のシートが既に存在します: " & newName, vbExclamation
                    errorCount = errorCount + 1
                    skipFlag = True
                End If
            End If
            
            If Not skipFlag Then
                ' シート名変更
                On Error Resume Next
                Set ws = ActiveWorkbook.Worksheets(oldName)
                ws.Name = newName
                
                If Err.Number = 0 Then
                    successCount = successCount + 1
                Else
                    errorCount = errorCount + 1
                    MsgBox "シート名の変更に失敗しました: " & oldName & " → " & newName & vbCrLf & _
                           "エラー: " & Err.description, vbExclamation
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    ' 結果報告
    MsgBox "処理完了" & vbCrLf & _
           "成功: " & successCount & " 件" & vbCrLf & _
           "失敗: " & errorCount & " 件", vbInformation
    
    ' シート一覧を更新
    Call LoadSheetList
    lstPreview.Clear
End Sub

Private Function IsValidSheetName(sheetName As String) As Boolean
    ' シート名の有効性をチェック
    Dim invalidChars As String
    Dim i As Integer
    
    invalidChars = "\/?*[]:"
    
    ' 空文字チェック
    If Trim(sheetName) = "" Then
        IsValidSheetName = False
        Exit Function
    End If
    
    ' 長さチェック（Excelは31文字まで）
    If Len(sheetName) > 31 Then
        IsValidSheetName = False
        Exit Function
    End If
    
    ' 無効文字チェック
    For i = 1 To Len(invalidChars)
        If InStr(sheetName, Mid(invalidChars, i, 1)) > 0 Then
            IsValidSheetName = False
            Exit Function
        End If
    Next i
    
    IsValidSheetName = True
End Function

Private Function sheetExists(sheetName As String) As Boolean
    ' 指定されたシートが存在するかチェック
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    sheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub btnCancel_Click()
    ' フォームを閉じる
    Unload Me
End Sub

' テキストボックス・コンボボックスの変更イベント
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

' 連番関連の変更イベント
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

' 選択されたシートのバックアップを作成
Sub CreateSheetBackup()
    Dim ws As Worksheet
    Dim backupWb As Workbook
    Dim originalName As String
    
    Set backupWb = Workbooks.Add
    originalName = ActiveWorkbook.Name
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Copy After:=backupWb.Sheets(backupWb.Sheets.count)
    Next ws
    
    ' 最初の空シートを削除
    Application.DisplayAlerts = False
    backupWb.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    backupWb.SaveAs ThisWorkbook.Path & "\シート名変更前バックアップ_" & _
                    Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    
    MsgBox "バックアップを作成しました: " & backupWb.Name, vbInformation
End Sub


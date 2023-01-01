
 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err
    
    '画面の更新をオフにする
    Application.ScreenUpdating = False
    
    'フォルダパス
    strFolderPath = ThisWorkbook.Worksheets(MAIN_SHEET_NAME).Cells(1, 1).Value
    'バックアップフォルダパス
    strBackUpFolderPath = ThisWorkbook.Worksheets(MAIN_SHEET_NAME).Cells(2, 1).Value
    '出力先フォルダパス
    strOutputFolderPath = ThisWorkbook.Worksheets(MAIN_SHEET_NAME).Cells(3, 1).Value
    
    'TODO : 改善の余地あり
    '入力チェック
    If ExecFunc(Func.IsFolderPath, strFolderPath) = False Then
        Call MsgBox(FOLDER_PATH_INVALID, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
    '入力チェック
    If ExecFunc(Func.IsFolderPath, strBackUpFolderPath) = False Then
        Call MsgBox(FOLDER_PATH_INVALID, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
    '入力チェック
    If ExecFunc(Func.IsFolderPath, strOutputFolderPath) = False Then
        Call MsgBox(FOLDER_PATH_INVALID, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
    '処理を実行する
    If ExecPerExtension(ThisWorkbook.Path, strBackUpFolderPath, DATA_SHEET_NAME) = False Then
        Call MsgBox(PROCESS_FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
    '画面の更新をオンにする
    Application.ScreenUpdating = True
End Sub


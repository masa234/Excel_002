



'【概要】txtファイル群をExcelのファイルとして出力する
Public Function TxtFilesToExceFile(ByVal txtFilePaths As Variant, _
                                ByVal strBaseSheetName As String) As Boolean
On Error GoTo TxtFilesToExceFile_Err

    TxtFilesToExceFile = False
    
    Dim lngArrIdx As Long
    Dim strTxtFilePath As String
    Dim objWb As Excel.Workbook
        
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(txtFilePaths)
        'テキストファイルパス
        strTxtFilePath = txtFilePaths(lngArrIdx)
        'Excelファイル作成
        Set objWb = Workbooks.Add
        'シート名を設定する
        strSheetName = GetSheetNameWithSeqNumber(objWb, strBaseSheetName)
        'Excelのシートとして展開
        If TxtFileToExcelSheet(strTxtFilePath, objWb, strSheetName) = False Then
            GoTo TxtFilesToExceFile_Exit
        End If
    Next lngArrIdx
    
    TxtFilesToExceFile = True
    
TxtFilesToExceFile_Err:

TxtFilesToExceFile_Exit:
    Set objWb = Nothing
End Function


'【概要】連番付きシート名を取得
Public Function GetSheetNameWithSeqNumber(ByVal objWb As Excel.Workbook, _
                                ByVal strBaseSheetName As String) As String
On Error GoTo GetSheetNameWithSeqNumber_Err

    Dim lngCount As Long
    Dim strSheetName As String

    '100回繰り返す
    For lngCount = 1 To 100
        'ファイル名設定
        strSheetName = strBaseSheetName & "_" & CStr(lngCount)
        'シートが存在しない場合、処理終了
        If IsExistsSheet(objWb, strSheetName) = False Then
            GetSheetNameWithSeqNumber = strSheetName
            GoTo GetSheetNameWithSeqNumber_Exit
        End If
    Next lngCount
    
GetSheetNameWithSeqNumber_Err:

GetSheetNameWithSeqNumber_Exit:

End Function


'【概要】ファイル名日付付きを取得
Public Function GetFileNameWithDate(ByVal strDirectoryPath As String, _
                                ByVal strBaseFileName As String) As Boolean
On Error GoTo GetFileNameWithDate_Err

    GetFileNameWithDate = False

    Dim lngCount As Long
    Dim strFileName As String
    Dim strFilePath As String

    '100回繰り返す
    For lngCount = 1 To 100
        'ファイル名設定
        strFileName = strBaseFileName & "_" & Format(yyyy_mm_dd) & CStr(lngCount)
        'ファイルパス
        strFilePath = strDirectoryPath & "\" & strFileName & ".xlsx"
        'ファイルが存在しない場合、処理終了
        If ExecFunc(Func.IsExistsFile, strFilePath) = False Then
            GetFileNameWithDate = strFileName
            GoTo GetFileNameWithDate_Exit
        End If
    Next lngCount
    
    GetFileNameWithDate = True
    
GetFileNameWithDate_Err:

GetFileNameWithDate_Exit:

End Function



'【概要】ファンクションを実行する
Public Function ExecFunc(ByVal enumFuncType As String, _
                        Optional ByVal strFilePath As String) As Boolean
On Error GoTo ExecFunc_Err

    ExecFunc = False

    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    'それぞれ違う処理を適応する
    Select Case enumFuncType
        Case Func.IsExistsFile
            'ファイルが存在する場合、True
            If objFso.FileExists(strFilePath) = True Then
                ExecFunc = True
                GoTo ExecFunc_Exit
            End If
        Case Func.DeleteFile
            'ファイル削除
            Kill strFilePath
            '削除に成功した場合、True
            ExecFunc = True
        Case Func.IsFolderPath
            'フォルダが存在する場合、True
            If objFso.FolderExists(strFilePath) = True Then
                ExecFunc = True
                GoTo ExecFunc_Exit
            End If
    End Select
    
ExecFunc_Err:

ExecFunc_Exit:
    Set objFso = Nothing
End Function


'【概要】拡張子毎に処理を実行する
Public Function ExecPerExtension(ByVal strDirectoryPath As String, _
                            ByVal strBackUpFolderPath As String, _
                            ByVal strSheetName As String) As Boolean
On Error GoTo ExecPerExtension_Err

    ExecPerExtension = False
    
    Dim lngArrIdx As Long
    Dim strFilePath As String
    Dim strOutputSheetName As String
    Dim strCopyFolderPath As String
    Dim strCopiedFolderPath As String
    Dim strExtensionName As String
    Dim arrFilePaths() As Variant
    Dim objWb As Excel.Workbook

    'ファイル群を取得する
    arrFilePaths = GetFilePaths(strDirectoryPath)
        
    'Excelのブック作成
    Set objWb = Workbooks.Add
    
    'シート名をDATAにする
    ActiveSheet.Name = strSheetName
    strOutputSheetName = strSheetName
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrFilePaths)
        'ファイルパス
        strFilePath = arrFilePaths(lngArrIdx)
        '拡張子ごとに処理を行う
        strExtensionName = GetExtensionName(strFilePath)
        Select Case strExtensionName
            '拡張子が.accdbの場合
            Case ACCESS_EXTENSION_NAME
                'コピー先フォルダパス
                strCopiedFolderPath = strBackUpFolderPath & "\" & GetFileName(strFilePath)
                'フォルダコピー
                FileCopy strFilePath, strCopiedFolderPath
                'ファイルを削除する
                If ExecFunc(Func.DeleteFile, strFilePath) = False Then
                    GoTo ExecPerExtension_Exit
                End If
            '拡張子が.txtの場合
            Case TXT_EXTENSION_NAME
                'シートに内容をコピー
                If TxtFileToExcelSheet(strFilePath, objWb, strOutputSheetName) = False Then
                    GoTo ExecPerExtension_Exit
                End If
                '次のシート名取得
                strOutputSheetName = GetSheetNameWithSeqNumber(objWb, strSheetName)
                '最後のループの場合は、シート作成しない
                If lngArrIdx <> UBound(arrFilePaths) Then
                    'シート作成
                    objWb.Worksheets.Add
                    'シート名変更
                    ActiveSheet.Name = strOutputSheetName
                End If
        End Select
    Next lngArrIdx

    ExecPerExtension = True
    
ExecPerExtension_Err:

ExecPerExtension_Exit:
    Set objFso = Nothing
End Function


'定数
Public Const MAIN_SHEET_NAME = "メイン"
Public Const DATA_SHEET_NAME = "Data"
Public Const SAVE_BOOK_NAME = "Excelデータ"
Public Const SAVE_TEXT_NAME = "データ"
Public Const CONFIRM = "確認"
Public Const VB_EXTENSION_NAME = "vb"
Public Const ACCESS_EXTENSION_NAME = "accdb"
Public Const TXT_EXTENSION_NAME = "txt"

'メッセージ
Public Const FOLDER_PATH_INVALID = "正しいフォルダパスを指定してください。"
Public Const PROCESS_FAILED = "処理に失敗しました。"

'フォルダパス
Public strFolderPath As String
'バックアップフォルダパス
Public strBackUpFolderPath As String
'出力先フォルダパス
Public strOutputFolderPath As String


'関数
Public Enum Func
    IsExistsFile = 1
    IsFolderPath = 2
    DeleteFile = 3
End Enum

'【概要】txtファイルをExcelのシートとして出力する
Public Function TxtFileToExcelSheet(ByVal txtFilePath As String, _
                                ByVal objWb As Excel.Workbook, _
                                ByVal strSheetName As String) As Boolean
On Error GoTo TxtFileToExcelSheet_Err

    TxtFileToExcelSheet = False
    
    Dim lngFreeFile As Long
    Dim strLine As String

    'フリーファイルを取得する
    lngFreeFile = FreeFile
    
    'テキストファイルを開く
    Open txtFilePath For Input As #lngFreeFile
    
    '列初期化
    lngCurrentRow = 1
    
    '終端まで繰り返す
    Do Until EOF(lngFreeFile)
        '1行読み込み
        Line Input #lngFreeFile, strLine
        '出力
        objWb.Worksheets(strSheetName).Cells(lngCurrentRow, 1).Value = strLine
        '行をカウントアップ
        lngCurrentRow = lngCurrentRow + 1
    Loop
    
    TxtFileToExcelSheet = True
    
TxtFileToExcelSheet_Err:

TxtFileToExcelSheet_Exit:
    'テキストファイルを開く
    Close #lngFreeFile
End Function


'【概要】シートが存在するか
Public Function IsExistsSheet(ByVal objWb As Excel.Workbook, _
                           ByVal strSheetName As String) As Boolean
On Error GoTo IsExistsSheet_Err

    IsExistsSheet = False
    
    Dim objWs As Excel.Worksheet
    
    '引数のシート名でシートオブジェクトを参照する
    'シートが存在しない場合、エラーが発生する
    Set objWs = objWb.Worksheets(strSheetName)
    
    IsExistsSheet = True
    
IsExistsSheet_Err:

IsExistsSheet_Exit:
    Set objWs = Nothing
End Function


'【概要】特定のディレクトリの特定の拡張子のファイル群を取得する
Public Function GetFilePaths(ByVal strDirectoryPath As String) As Variant
On Error GoTo GetFilePaths_Err
    
    Dim lngArrIdx As Long
    Dim arrRet() As Variant
    Dim objFso As FileSystemObject
    Dim objFile As File
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    With objFso
        'ファイルの数だけ繰り返す
        For Each objFile In .GetFolder(strDirectoryPath).Files
            '配列再宣言
            ReDim Preserve arrRet(lngArrIdx)
            '配列に格納
            arrRet(lngArrIdx) = objFile.Path
            '配列の要素番号を1つ進める
            lngArrIdx = lngArrIdx + 1
        Next objFile
    End With
        
    GetFilePaths = arrRet
    
GetFilePaths_Err:

GetFilePaths_Exit:
    Set objFso = Nothing
    Set objFile = Nothing
End Function


'【概要】拡張子を取得する
Public Function GetExtensionName(ByVal strFilePath As String) As String
On Error GoTo GetExtensionName_Err
    
    Dim objFso As FileSystemObject
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
        
    '拡張子を取得
    GetExtensionName = objFso.GetExtensionName(strFilePath)
    
GetExtensionName_Err:

GetExtensionName_Exit:
    Set objFso = Nothing
End Function


'★TODO：改善の余地あり
'【概要】ファイル名を取得する
Public Function GetFileName(ByVal strFilePath As String) As String
On Error GoTo GetFileName_Err
    
    Dim objFso As FileSystemObject
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
        
    '拡張子を取得
    GetFileName = objFso.GetFileName(strFilePath)
    
GetFileName_Err:

GetFileName_Exit:
    Set objFso = Nothing
End Function

Option Explicit

'=============================================================================
' モジュール名  : Gazou_harituke
' 作成目的      : 指定フォルダ配下のサブフォルダごとに画像（PNG/JPG）を
'                 Excelシートへ自動貼り付けするマクロ群
'
' 【前提】──────────────────────────────────────────────────────
'   ・対象フォルダの直下に「処理番号_担当_シート名」形式のサブフォルダが並んでいること
'     例）01_山田_申請受付、02_鈴木_審査、…
'     ※シート名はフォルダ名を「_」で分割した 3 番目の要素（index 2）を使用する
'   ・サブフォルダ内に PNG または JPG の画像ファイルが格納されていること
'   ・貼り付け対象ブックは本マクロが含まれるブック自身
'
' 【公開マクロ一覧】────────────────────────────────────────────
'
'  ■メイン
'   ImportImagesTool   … フォルダを選択してシート生成＋画像貼り付けを一括実行
'
' 【内部ヘルパー（直接実行は非推奨）】
'   SelectFolder       … フォルダ選択ダイアログを表示し選択パスを返す
'   SelectImageScale   … 画像サイズ倍率を選択させる（1 / 0.8 / 0.5）
'   GetSheetName       … フォルダ名からシート名部分を切り出す
'   ImportImages       … フォルダ内の PNG/JPG を順にシートへ貼り付ける
'   InsertImage        … 1 画像をシートへ貼り付け、次の挿入行番号を返す
'   DeleteSheetIfExists … 指定名のシートが存在すれば削除する
'   CreateIndexSheet   … INDEX シートを先頭に作成する
'   AddToIndex         … INDEX シートに各シートへのハイパーリンクを追加する
'
' 【使い方（基本フロー）】──────────────────────────────────────
'   1. ImportImagesTool を実行する
'   2. 画像が格納されたサブフォルダを含む親フォルダを選択する
'   3. 画像サイズ倍率を選択する（既定: 80%）
'   4. サブフォルダごとにシートが生成され、画像が貼り付けられる
'   5. 先頭に INDEX シートが作成され、各シートへのリンクが一覧表示される
'
' 【ファイル名の表示ルール】─────────────────────────────────────
'   ファイル名の先頭にあるタイムスタンプ（アンダースコア 3 つまでの部分）は
'   自動的に除去してシートに表示します。
'   例）20260403_162311_453_申請書.png → 申請書.png
'   アンダースコアが 3 つ未満のファイルはファイル名をそのまま表示します。
'   この変換は GetDisplayFileName 関数で行います。
'
' 【モジュール変数について】──────────────────────────────────────
'   IMAGE_SCALE      … 画像の表示倍率（SelectImageScale で設定し InsertImage で参照）
'   ENABLE_HYPERLINK … 画像クリック時のハイパーリンク設定有無（SelectHyperlinkOption で設定）
'
' 【変更履歴】──────────────────────────────────────────────────
'   v1.0  初版作成
'   v1.1  ハイパーリンク設定の選択機能追加 / ファイル名タイムスタンプ除去対応
'         実行前注意メッセージ（ShowStartupNotice）追加
'=============================================================================

' INDEX シート名の定数（複数箇所で使用するためモジュール先頭で一元管理）
Private Const INDEX_SHEET_NAME As String = "INDEX"

' 画像の表示倍率（SelectImageScale で設定）
Private IMAGE_SCALE As Double

' 画像クリック時のハイパーリンク設定有無（SelectHyperlinkOption で設定）
Private ENABLE_HYPERLINK As Boolean

Sub ImportImagesTool()

    ' 実行前にファイル名変換ルールとリンク設定についてユーザーへ案内
    ShowStartupNotice

    Dim basePath As String
    basePath = SelectFolder()

    If basePath = "" Then Exit Sub

    IMAGE_SCALE = SelectImageScale()
    ENABLE_HYPERLINK = SelectHyperlinkOption()

    ' エラー時も ScreenUpdating を必ず復元するためハンドラを設定
    On Error GoTo Cleanup

    Application.ScreenUpdating = False

    CreateIndexSheet

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As Object
    Set parentFolder = fso.GetFolder(basePath)

    ' ループ変数はスコープを明確にするためループ前に宣言
    Dim subFolder As Object
    Dim sheetName As String
    Dim ws As Worksheet

    For Each subFolder In parentFolder.SubFolders

        sheetName = GetSheetName(subFolder.name)

        DeleteSheetIfExists sheetName

        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.name = sheetName

        AddToIndex sheetName

        ImportImages ws, subFolder.path

    Next

Cleanup:
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
    Else
        MsgBox "完了"
    End If

End Sub


'-----------------------------------------------------------------------------
' フォルダ選択ダイアログを表示し、選択されたフォルダのパスを返す
' 戻り値 : 選択パス文字列（キャンセル時は空文字）
'-----------------------------------------------------------------------------
Function SelectFolder() As String

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "画像フォルダを選択"

        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        End If

    End With

End Function


'-----------------------------------------------------------------------------
' 画像サイズ倍率を InputBox で選択させる
' 戻り値 : 1.0（100%）/ 0.8（80%）/ 0.5（50%）のいずれか（既定: 0.8）
'-----------------------------------------------------------------------------
Function SelectImageScale() As Double

    Dim ans As String

    ans = InputBox( _
    "画像サイズを選択" & vbCrLf & _
    "1 = 100%" & vbCrLf & _
    "2 = 80%" & vbCrLf & _
    "3 = 50%", _
    "画像サイズ", "2")

    Select Case ans
        Case "1"
            SelectImageScale = 1
        Case "2"
            SelectImageScale = 0.8
        Case "3"
            SelectImageScale = 0.5
        Case Else
            SelectImageScale = 0.8
    End Select

End Function


'-----------------------------------------------------------------------------
' 実行前の注意事項をユーザーに表示する
' ファイル名の変換ルールとハイパーリンク設定について事前に案内する
'-----------------------------------------------------------------------------
Sub ShowStartupNotice()

    MsgBox "【実行前のご確認】" & vbCrLf & vbCrLf & _
           "■ ファイル名の表示について" & vbCrLf & _
           "  ファイル名先頭のタイムスタンプ（アンダースコア 3 つまでの部分）は" & vbCrLf & _
           "  自動的に除去してシートに表示します。" & vbCrLf & _
           "  例）20260403_162311_453_申請書.png  →  申請書.png" & vbCrLf & _
           "  ※アンダースコアが 3 つ未満の場合はファイル名をそのまま表示します。" & vbCrLf & vbCrLf & _
           "■ ハイパーリンクについて" & vbCrLf & _
           "  次の画面で画像クリック時のリンク設定（あり／なし）を選択できます。" & vbCrLf & vbCrLf & _
           "続けて各設定を選択してください。", _
           vbInformation, "画像貼り付けツール"

End Sub


'-----------------------------------------------------------------------------
' 画像クリック時のハイパーリンクを設定するかどうかをユーザーに選択させる
' 戻り値 : True（リンクあり）/ False（リンクなし）
'-----------------------------------------------------------------------------
Function SelectHyperlinkOption() As Boolean

    Dim ans As String

    ans = InputBox( _
        "画像クリック時のハイパーリンク設定を選択してください。" & vbCrLf & vbCrLf & _
        "1 = あり（画像クリックで元ファイルを開く）" & vbCrLf & _
        "2 = なし", _
        "ハイパーリンク設定", "1")

    SelectHyperlinkOption = (ans = "1")

End Function


'-----------------------------------------------------------------------------
' フォルダ名をアンダースコアで分割し、3 番目の要素をシート名として返す
' 例）"01_山田_申請受付" → "申請受付"
' 3 要素未満の場合はフォルダ名をそのまま返す
' Excel シート名の上限（31 文字）を超える場合は切り詰める
'-----------------------------------------------------------------------------
Function GetSheetName(folderName As String) As String

    Dim parts
    parts = Split(folderName, "_")

    Dim name As String
    If UBound(parts) >= 2 Then
        name = parts(2)
    Else
        name = folderName
    End If

    ' Excel シート名は 31 文字まで（超過するとエラーになるため切り詰め）
    GetSheetName = Left(name, 31)

End Function


'-----------------------------------------------------------------------------
' ファイル名先頭のタイムスタンプ部分を除去して表示用ファイル名を返す
' アンダースコア 3 つより後ろの文字列を返す
' 例）20260403_162311_453_申請書.png → 申請書.png
' アンダースコアが 3 つ未満の場合はファイル名をそのまま返す
'-----------------------------------------------------------------------------
Function GetDisplayFileName(file As String) As String

    Dim parts
    parts = Split(file, "_")

    ' アンダースコアが 3 つ以上ある（要素数が 4 以上）場合のみ変換する
    If UBound(parts) >= 3 Then

        Dim i As Long
        Dim result As String
        result = parts(3)

        For i = 4 To UBound(parts)
            result = result & "_" & parts(i)
        Next i

        GetDisplayFileName = result

    Else

        GetDisplayFileName = file

    End If

End Function


'-----------------------------------------------------------------------------
' 指定フォルダ内の PNG/JPG ファイルをファイル名順に ws へ貼り付ける
' Dir で全ファイルを走査し、拡張子で PNG/JPG のみを対象とする
' ※拡張子ごとに 2 パスで処理すると混在フォルダでファイル名順が崩れるため
'   1 パスで全ファイルを走査して拡張子を判定する方式を採用する
'-----------------------------------------------------------------------------
Sub ImportImages(ws As Worksheet, folderPath As String)

    Dim file As String
    Dim row As Long
    Dim ext As String

    row = 1

    file = Dir(folderPath & "\*.*")

    Do While file <> ""

        ext = LCase(Right(file, 3))

        If ext = "png" Or ext = "jpg" Then
            row = InsertImage(ws, folderPath, file, row)
        End If

        file = Dir

    Loop

End Sub


'-----------------------------------------------------------------------------
' 画像 1 枚をシートへ貼り付け、次の挿入開始行番号を返す
' 引数 : ws    … 貼り付け先シート
'        path  … 画像ファイルのフォルダパス
'        file  … ファイル名（拡張子含む）
'        row   … 現在の挿入開始行番号
' 戻り値 : 次の画像の挿入開始行番号
'   ・B 列（row 行）にファイル名を記載し、画像は C 列（row+1 行）に配置する
'   ・画像の高さをもとに次の挿入行を算出し、間に 3 行のマージンを設ける
'-----------------------------------------------------------------------------
Function InsertImage(ws As Worksheet, path As String, file As String, row As Long) As Long

    ' B 列にタイムスタンプを除去した表示用ファイル名を記載
    ws.Cells(row, 2).Value = GetDisplayFileName(file)

    Dim fullPath As String
    fullPath = path & "\" & file

    ' 画像の配置基準セル（ファイル名行の 1 行下・C 列）
    Dim targetCell As Range
    Set targetCell = ws.Cells(row + 1, 3)

    Dim pic As Shape

    ' Width/Height に -1 を指定すると元のサイズで読み込まれる
    Set pic = ws.Shapes.AddPicture( _
        Filename:=fullPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=-1, _
        Height:=-1)

    With pic

        ' アスペクト比を維持したまま Width のみ倍率変換する
        .LockAspectRatio = msoTrue
        .Width = .Width * IMAGE_SCALE

        ' AddPicture 直後は位置がずれることがあるため C 列に固定し直す
        .Left = targetCell.Left
        .Top = targetCell.Top

        ' セルの移動・サイズ変更に追随させる
        .Placement = xlMoveAndSize

    End With

    ' ユーザー選択に応じて画像クリック時のハイパーリンクを設定する
    If ENABLE_HYPERLINK Then
        ws.Hyperlinks.Add _
            Anchor:=pic, _
            Address:=fullPath
    End If

    ' 画像の高さ（pt）を標準行高さで割って使用行数を算出し、3 行のマージンを加算
    Dim rowsUsed As Long
    rowsUsed = pic.Height / ws.Rows(1).RowHeight + 3

    InsertImage = row + rowsUsed

End Function


'-----------------------------------------------------------------------------
' 指定名のシートが存在すれば削除する（存在しない場合は何もしない）
' 確認ダイアログは DisplayAlerts で抑制する（エラー時も必ず復元）
'-----------------------------------------------------------------------------
Sub DeleteSheetIfExists(name As String)

    Dim ws As Worksheet

    For Each ws In Worksheets

        If ws.name = name Then

            Application.DisplayAlerts = False
            On Error GoTo RestoreAlerts
            ws.Delete
RestoreAlerts:
            Application.DisplayAlerts = True

            If Err.Number <> 0 Then Err.Raise Err.Number

            Exit Sub

        End If

    Next

End Sub


'-----------------------------------------------------------------------------
' INDEX シートを先頭（Sheets(1) の前）に新規作成する
' 既存の INDEX シートは事前に削除してから作成する
'-----------------------------------------------------------------------------
Sub CreateIndexSheet()

    DeleteSheetIfExists INDEX_SHEET_NAME

    Dim ws As Worksheet
    Set ws = Sheets.Add(Before:=Sheets(1))

    ws.name = INDEX_SHEET_NAME

    ws.Cells(1, 1) = "Sheet List"

End Sub


'-----------------------------------------------------------------------------
' INDEX シートの末尾行に、指定シートへのハイパーリンクを追加する
'-----------------------------------------------------------------------------
Sub AddToIndex(sheetName As String)

    Dim ws As Worksheet
    Set ws = Sheets(INDEX_SHEET_NAME)

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1

    ws.Hyperlinks.Add _
        Anchor:=ws.Cells(r, 1), _
        Address:="", _
        SubAddress:="'" & sheetName & "'!A1", _
        TextToDisplay:=sheetName

End Sub


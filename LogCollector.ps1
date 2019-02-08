<#
    .SYNOPSIS
        イベント ログ保存スクリプト "LogCollector.ps1" ver.0.21.20190208 山岸真人@株式会社ロビンソン
     
    .DESCRIPTION
        Windows が出力したイベント ログを指定した期間、指定したフォルダーに保存します。
     
    .PARAMETER Events
        収集するイベント ログの種類を指定します。
        Application、Security、System のいずれかもしくは複数を指定できます。
        省略した場合はイベント ログを収集しません。

    .PARAMETER SourceFolderPath
        収集するファイルをフルパスもしくは相対パスで指定します。
        複数指定する場合はカンマで区切ります。
        スペースを含む場合は "Full Path" もしくは 'Full Path' のように指定します。
        省略した場合はファイルを収集しません。
        この機能は現在開発中です。
    
    .PARAMETER StartDate
        収集対象ログの範囲を DateTime 型で指定します。
        このパラメーターで指定した日時と同じかそれより新しいログが収集されます。
        省略した場合は前日の 0 時 0 分 0 秒が指定されます。
     
    .PARAMETER EndDate
        収集対象ログの範囲を DateTime 型で指定します。
        このパラメーターで指定した日時より古いログが収集されます。
        省略した場合は StartDate 引数の 1 日後が指定されます。

    .PARAMETER ZipPath
        収集したログの圧縮ファイルを保存するフォルダーをフルパスもしくは相対パスで指定します。
        このパラメーターは省略できません。

    .PARAMETER ZipLifeTime
        作成した圧縮ファイルの保存期間を日数で指定します。
        保存期間を超過した圧縮ファイルは削除されます。
        保存期間の時分秒は切り捨てられます。
        省略した場合、ゼロもしくはマイナスの値が指定された場合は圧縮ファイルを削除しません。

    .INPUTS
        なし。パイプラインからの入力には対応していません。
     
    .OUTPUTS
        正常終了もしくは警告終了した場合は 0 を、異常終了した場合は 1 を返します。
     
    .EXAMPLE
        .\LogCollector.ps1 -Events System,Application -ZipPath C:\Temp\Log
        C:\Temp\Log に System、Application イベント ログをエクスポートします。
        -StartDate および -EndDate が指定されていないため、前日 1 日分が収集対象となります。

    .NOTES
        Windows 10 にて動作確認を実施しています。
        Windows 10 および Windows Server 2016 以降に搭載された PowerShell をサポートします。
        不具合のご連絡は以下の Github リポジトリまでどうぞ。
        https://github.com/yamanyon/LogCollector
        有償サポートのご依頼は m.yamagishi@robinsons.co.jp もしくは twitter: yamanyon までどうぞ。
#>

<###################################################################################################
v0.20 2018/12/20 初期リリース
v0.21 2019/02/08 ログ収集部分を改良, 主処理系を関数化, コメント改修
###################################################################################################>

<###################################################################################################
スクリプト引数定義
###################################################################################################>
param (
    # どのイベントログを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は取得しない
    [ValidateSet("Application","Security","System")]
    [array]$Events,
    
    # どのファイルを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は取得しない
    [array]$SourceFolderPath,

    # いつからのログを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は前日の 0:00:00 を設定
    [DateTime]$StartDate = (Get-Date).Date.AddDays(-1),

    # いつまでのログを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は -StartDate に 1 日加算
    [DateTime]$EndDate = $StartDate.AddDays(1),

    # ログの保存先を指定する
    [parameter(mandatory=$true)]    # null を許容しない
    [String]$ZipPath,

    # 保存先ログの保存期間を日数で指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は無限に保存する
    [int]$ZipLifeTime
)

<###################################################################################################
スクリプト環境変数定義
###################################################################################################>
# PowerShell Cmdlet のエラー発生時の処理を停止にする
$ErrorActionPreference = "Stop"

<###################################################################################################
主処理
###################################################################################################>
function LogCollector {
    WriteLog "Info" "処理を開始します。"

    # ふたつの日付引数のチェック
    funcDateCheck
    
    # 指定された -ZipPath パラメータからコピー先フォルダーのフルパスを取得
    $strDestinationNowFolder = CreateDestFolder $ZipPath
    
    # イベント ログの出力処理
    funcEventLog $strDestinationNowFolder $Events
    
    # ファイル (所謂テキスト形式ログ) の収集
    funcSourceFolders $strDestinationNowFolder $SourceFolderPath
    
    # ZIP 圧縮処理
    funcCompress $strDestinationNowFolder $ZipPath
    
    # 圧縮後の収集ログ削除処理
    DeleteFolder $strDestinationNowFolder
    
    # 保存期間を超過した圧縮ファイルを割くぞする
    DeleteExpiredFile $ZipPath $ZipLifeTime
    
    WriteLog "Info" "すべての処理が完了しました。"
    exit 0
}

<###################################################################################################
ログを記録する
戻り値: なし
###################################################################################################>
function WriteLog {
    param (
        [ValidateSet("Info","Warn","Error")]
        [String]$strType,   # メッセージのタイプ
        [String]$strMessage # メッセージ文字列
    )

    # 記録日時の文字列を生成
    $strNow = ( Get-Date ).ToString("yyyy/MM/dd HH:mm:ss")

    # Standard I/O に出力
    Write-Host "$strNow [$strType] $strMessage"

    # Warn か Error が指定されたら $Error[0] も出力
    if ( $strType -eq "Warn" -or $strType -eq "Error" ) {
        Write-Host $Error[0]
    }

    # このあたりにイベント ログへのログ出力機能も追加したい

    # $strType に Error が指定されたらエラー終了
    if ( $strType -eq "Error" ) {
        exit 1
    }
}

<###################################################################################################
収集開始引数と収集終了日時引数の比較
戻り値: なし
###################################################################################################>
function funcDateCheck {
    try {
        if ( $StartDate -ge $EndDate ) {
            throw ( "-EndDate ( " + $EndDate.ToString("yyyy/MM/dd HH:mm:ss") + " ) より -StartDate ( " + $StartDate.ToString("yyyy/MM/dd HH:mm:ss") + " ) が古くなくてはいけません。" )
        }
    }
    catch {
        WriteLog "Error" "funcDateCheck error."
    }

    WriteLog "Info" ( $StartDate.ToString("yyyy/MM/dd HH:mm:ss") + " ～ " + $EndDate.ToString("yyyy/MM/dd HH:mm:ss") + " までのイベント ログやファイルを出力します。" )
}

<###################################################################################################
指定されたフォルダーパスが正常にアクセスできるか確認し、正常ならフォルダーオブジェクトを返す
戻り値: $objFolder 指定されたパスで取得できたフォルダオブジェクト
###################################################################################################>
function GetFolderObject {
    param (
        [String]$strFolder  # フォルダーのパス
    )

    # フォルダー文字列が利用できるかどうか一旦確認する
    try {
        $bTestResult = Test-Path $strFolder
    }
    catch {
        WriteLog "Error" "GetFolderObject $strFolder is invalid string."
    }

    # フォルダー 文字列が利用できなければ Test-path の結果に応じて作成/取得をしてみる。フォルダーを試しに作成してみる
    if ( $bTestResult ) {
        # 指定された文字列で Get-Item してみる
        try {
            $objFolder = Get-Item $strFolder
        }
        catch {
            WriteLog "Error" "GetFolderObject Get-Item $strFolder cannot access."
        }

        # Get-Item できたらそれがコンテナオブジェクトか確認する
        try {
            if ( !$objFolder.PSIsContainer ) {
                throw
            }
        }
        catch {
            WriteLog "Error" "GetFolderObject $strFolder is not container."
        }
    }
    else {
        $objFolder = CreateFolder $strFolder
    }

    # Get-Item できてかつコンテナならそのオブジェクトを返す
    $objFolder
}

<###################################################################################################
指定されたフォルダーパスを作成する
戻り値: $objFolder 指定されたパスで取得できたフォルダオブジェクト
###################################################################################################>
function CreateFolder {
    param (
        [String]$strFolder  # フォルダーのパス
    )

    # 指定された文字列で New-Item してみる
    try {
        $objFolder = New-Item $strFolder -Type Directory
    }
    catch {
        WriteLog "Error" "CreateFolder New-Item $strFolder cannot complete."
    }

    # Get-Item できてかつコンテナならそのオブジェクトを返す
    $objFolder
}

<###################################################################################################
イベントオブジェクトを取得する
戻り値: $objFolder 指定されたパスで取得できたフォルダオブジェクト
###################################################################################################>
function CreateDestinationNowFolder {
    param (
        [String]$strDestination # フォルダ文字列
    )
    # フォルダ名文字列を生成する
    $strPath  = $strDestination + "\"
    $strPath += ( Get-Date ).ToString("yyyyMMddHHmmss")

    # フォルダを作成
    $objFolder = CreateFolder $strPath

    # オブジェクトを返す
    $objFolder
}

<###################################################################################################
イベントオブジェクトをCSVにエクスポートする
戻り値: なし
###################################################################################################>
function GetEventObjects {
    param (
        [String]$strEventLog,   # 取得するイベントログの名前
        [Object]$strWriteFolder # 保存先フォルダーのパス
    )

    # 指定されたイベントログを取得し $objEvents に保存する
    try {
        $objEvents = Get-EventLog $strEventLog `
        |   Where-Object { $_.TimeGenerated -ge $StartDate } `
        |   Where-Object { $_.TimeGenerated -lt $EndDate }
    }
    catch {
        WriteLog "Warn" "GetEventObjects Get-EventLog $strEventLog export error, but the script is contiunue."
        return
    }

    # 日付型を日付文字列に変換するための文字列を定義する
    $strFormat = "yyyyMMddHHmmss"

    # ファイル名文字列を生成する
    $strPath  = $strWriteFolder + "\"
    $strPath += $strEventLog + "_"
    $strPath += $StartDate.ToString($strFormat) + "_"
    $strPath += $EndDate.ToString($strFormat) + ".csv"

    # ファイル名を指定して CSV で出力する
    WriteLog "Info" "$strPath にイベント ログ $strEventLog をエクスポートします。"
    try {
        $objEvents | Export-Csv -Path $strPath -Encoding Default -NoTypeInformation
    }
    catch {
        WriteLog "Error" "GetEventObjects Export-Csv $strPath Error."
    }
}

<###################################################################################################
収集開始引数と収集終了日時引数の比較
戻り値: $strDestinationNowFolder ログの保存先フォルダ文字列
###################################################################################################>
function CreateDestFolder {
    param (
        [string]$strFolder  # 作成するフォルダの名前
    )
    WriteLog "Info" "フォルダー $strFolder が正常に利用できるか確認します。"
    getFolderObject $strFolder | Out-Null

    WriteLog "Info" "フォルダー $strFolder にログを保存するためのフォルダーを作成します。"
    $objDestinationNowFolder = CreateDestinationNowFolder $strFolder
    $strDestinationNowFolder = $objDestinationNowFolder.FullName
    WriteLog "Info" "作成したフォルダー: $strDestinationNowFolder"

    $strDestinationNowFolder
}

<###################################################################################################
イベントログ系の処理
戻り値: なし
###################################################################################################>
function funcEventLog {
    param (
        $strDestinationNowFolder,   # ログを保存するフォルダー
        $strEvents                  # イベント ログ
    )

    if ( $null -eq $strEvents ) {
        WriteLog "Info" "-Events 引数が指定されていないので、イベント ログの収集処理は実行しません。"
        return
    }

    # イベント ログを保存するフォルダーを作成する
    WriteLog "Info" "$strDestinationNowFolder フォルダーにイベント ログを保存するためのフォルダーを作成します。"
    $objWriteFolder = CreateFolder "$strDestinationNowFolder\EventLog"
    $strWriteFolder = $objWriteFolder.FullName
    WriteLog "Info" "作成したフォルダー: $strWriteFolder"

    WriteLog "Info" "イベントログの収集を開始します。"
    foreach ( $strEventLog in $strEvents ) {
        GetEventObjects $strEventLog $strWriteFolder
    }
}

<###################################################################################################
ファイルオブジェクトの収集処理
戻り値: $objChildItems: 収集対象のファイル
        $null の場合対象なし
###################################################################################################>
function GetFileObjects {
    param (
        [string]$strPath    # 収集対象フォルダのフルパスもしくは相対パス
    )

    # $strPath がアクセスできるか確認
    try {
        $objPath = Get-Item $strPath
    }
    catch {
        WriteLog "Warn" "フォルダー $strPath へアクセスできません。ファイル収集処理をスキップします。"
        return $null
    }

    # 以降の処理でフルパスが利用しやすいように変数投入
    $strPathName = $objPath.FullName

    WriteLog "Info" "$strPathName フォルダーのファイルを収集します。"

    # ファイル一覧オブジェクトを収集する
    try {
        # v1.01 にて再帰処理化
        $objChildItems = Get-ChildItem $strPath -Recurse `
        |   Where-Object { $_.LastWriteTime -ge $StartDate } `
        |   Where-Object { $_.LastWriteTime -lt $EndDate }
    }
    catch {
        WriteLog "Warn" "フォルダー $strPathName に保存されいているファイルへアクセスできません。ファイル収集処理をスキップします。"
        return $null
    }

    # 収集した結果対象が無い場合の処理
    try {
        if ( $null -eq $objChildItems ) {
            throw "フォルダー $strPathName には収集対象となるファイルがありません。ファイル収集処理をスキップします。"
        }
    }
    catch {
        WriteLog "Warn" "フォルダー $strPathName には収集対象となるファイルがありません。ファイル収集処理をスキップします。"
        return $null
    }

    # 対象一覧を戻す
    return $objChildItems
}

<###################################################################################################
コピー先フォルダー作成処理
戻り値: $strDistFullPath (コピー先フルパス)
###################################################################################################>
function CreateDistFullpath {
    param (
        $strDistination,        # 収集先フォルダのパス
        $strSourceFolderPath    # 収集対象ファイルのパス
    )

    # $strSourceFolderPath がアクセスできるか確認
    try {
        $objPath = Get-Item $strSourceFolderPath
    }
    catch {
        WriteLog "Warn" "フォルダー $strSourceFolderPath へアクセスできません。コピー先フォルダパスが生成できません。"
        return $null
    }

    # ファイルを保存するフォルダーを作成する
    WriteLog "Info" "$strDestination フォルダーにファイルを保存するためのフォルダーを作成します。"
    $objWriteFolder = CreateFolder ( "$strDistination\" + $objPath.Name )
    $strDistFullPath = $objWriteFolder.FullName
    WriteLog "Info" "作成したフォルダー: $strDistFullPath"

    return $strDistFullPath
}

<###################################################################################################
ファイルコピー処理
戻り値: なし
###################################################################################################>
function CopyFiles {
    param (
        [object]$objChildItems, # コピー対象のファイル
        [string]$strDistFullPath # コピー先フォルダのフルパス
    )

    try {
        $objChildItems | Copy-Item -Destination $strDistFullPath
    }
    catch {
        WriteLog "Warn" "フォルダー $strDistFullPath にファイルをコピーできません。ファイルのコピー処理をスキップします。"
        return
    }
    
}

<###################################################################################################
ファイル収集処理
戻り値: なし
###################################################################################################>
function funcSourceFolders {
    param (
        [string]$strDestination,        # 収集先フォルダのフルパス
        [array]$strSourceFolderPaths    # 収集対象ファイルのフルパス
    )

    if ( $null -eq $strSourceFolderPaths ) {
        WriteLog "Info" "-SourceFolderPath 引数が指定されていないので、ファイルの収集処理は実行しません。"
        return
    }

    WriteLog "Info" "ファイルの収集を開始します。"
    foreach ( $strSourceFolderPath in $strSourceFolderPaths ) { 
        $objFiles = GetFileObjects $strSourceFolderPath
        if ( $null -ne $objFiles ) {
            $strDistFullPath = CreateDistFullpath $strDistination $strSourceFolderPath
            CopyFiles $objFiles $strDistFullPath
            }
    }
}

<###################################################################################################
ZIP 圧縮処理
戻り値: なし
###################################################################################################>
function funcCompress {
    param (
        [String]$strSource,     # 圧縮対象フォルダ
        [String]$strDestination # 圧縮ファイル保存先フォルダ
    )

    # $strSource がフォルダーかどうか確認しオブジェクトを生成
    $objSource = getFolderObject $strSource

    # $strSource 配下のファイルオブジェクトを収集
    try {
        $objSourceFiles = Get-ChildItem $objSource -Recurse | Where-Object { !$_.PSIsContainer }
    }
    catch {
        WriteLog "Warn" "funcCompress $strSource cannot access. Escape compress function."
        return
    }

    # $objSource フォルダーに保存されているオブジェクトにファイルがあることを確認
    try {
        if ( $null -eq $objSourceFiles ) {
            throw "funcCompress $strSource is null. Escape compress function."
        }
    }
    catch {
        WriteLog "Warn" "funcCompress $strSource is null. Escape compress function."
        return
    }
        
    # $objSource フォルダーに保存されているオブジェクトを取得
    try {
        $objSourceChileItems = Get-ChildItem $objSource
    }
    catch {
        WriteLog "Error" "funcCompress Get-ChildItem $strSource error."
    }

    # 取得できた $objSourceChildItems を String 型配列に放り込む
    $strSourceChildItems = $objSourceChileItems | ForEach-Object { $_.FullName }

    # 圧縮ファイルを保存するファイル名をフルパスで生成
    $strDestFileFullPath = $strDestination + "\" + $objSource.Name + ".zip"

    WriteLog "Info" "収集したログ ファイルを $strDestFileFullPath に圧縮します。"
    # 圧縮処理を実行
    try {
        Compress-Archive -Path $strSourceChildItems -DestinationPath $strDestFileFullPath -Force
    }
    catch {
        WriteLog "Error" "funcCompress Compress-Archive error."
    }
}

<###################################################################################################
フォルダーの削除
戻り値: なし
###################################################################################################>
function DeleteFolder {
    param (
        [String]$strFolderPath  # 削除対象フォルダーのパス
    )

    WriteLog "Info" "$strFolderPath フォルダーを削除します。"
    try {
        Remove-Item -Path $strFolderPath -Recurse -Force
    }
    catch {
        WriteLog "Error" "DeleteFolder $strFolderPath delete error."
    }
}

<###################################################################################################
圧縮ファイルの保存期間を DateTime 型に変換する
戻り値: $dtLifeTime 圧縮ファイルの保存期間
###################################################################################################>
function ConvertDateToDateTime {
    param (
        [int]$iDate # 何日前なのかを示す数値
    )

    # -ZipLifeTime で指定された保管日数を評価する。
    if ( $iDate -le 0 ) {
        # 0 以下の場合、この関数を $null で抜ける。
        return $null
    }

    # 保管日の閾値を算出する
    ( Get-Date ).AddDays($iDate*-1).Date
}

<###################################################################################################
保存期間が経過したファイルを削除する
戻り値: なし
###################################################################################################>
function DeleteExpiredFile {
    param (
        [string]$strFolder, # 削除対象フォルダー
        [int]$iLifeTime     # 保管日数
    )

    WriteLog "Info" "フォルダー $strFolder に保存されている、保存期間が経過したファイルを削除します。"

    $dtLifeTime = ConvertDateToDateTime $iLifeTime

    if ( $null -eq $dtLifeTime ) {
        WriteLog "Info" "-ZipLifeTime が未指定もしくは 0 以下のため、$strFolder のファイルは削除しません。"
        return
    }

    WriteLog "Info" ( $dtLifeTime.ToString("yyyy/MM/dd HH:mm:ss") + " より古いファイルを削除します。" )

    try {
        $objItems = Get-ChildItem $strFolder | Where-Object { $_.LastWriteTime -lt $dtLifeTime -and !$_.PSIsContainer }
    }
    catch {
        WriteLog "Error" "DeleteExpiredFile cannot get period expired file."
    }

    try {
        $objItems | ForEach-Object { Remove-Item $_.FullName -Force }
    }
    catch {
        WriteLog "Error" "DeleteExpiredFile cannot remove period expired file."
    }
}

<###################################################################################################
主処理実行
###################################################################################################>
LogCollector

<#
    .SYNOPSIS
        イベント ログ保存スクリプト "LogCollector.ps1" ver.0.2.20181218 山岸真人@株式会社ロビンソン
     
    .DESCRIPTION
        Windows が出力したイベント ログを指定した期間、指定したフォルダーに保存します。
     
    .PARAMETER Events
        収集するイベント ログの種類を指定します。
        Application、Security、System のいずれかもしくは複数を指定できます。
        省略した場合はイベント ログを収集しません。
    
    .PARAMETER StartDate
        収集対象ログの範囲を DateTime 型で指定します。
        このパラメーターで指定した日時と同じかそれより新しいログが収集されます。
        省略した場合は前日の 0 時 0 分 0 秒が指定されます。
     
    .PARAMETER EndDate
        収集対象ログの範囲を DateTime 型で指定します。
        このパラメーターで指定した日時より古いログが収集されます。
        省略した場合は StartDate 引数の 1 日後が指定されます。

    .PARAMETER DestPath
        収集したログの圧縮ファイルを保存するフォルダーをフルパスで指定します。
        このパラメーターは省略できません。

    .PARAMETER DestLifeTime
        作成した圧縮ファイルの保存期間を日数で指定します。
        保存期間を超過した圧縮ファイルは削除されます。
        省略した場合は圧縮ファイルを削除しません。
        現在この機能は開発中です。

    .INPUTS
        なし。パイプラインからの入力には対応していません。
     
    .OUTPUTS
        正常終了もしくは警告終了した場合は 0 を、異常終了した場合は 1 を返します。
     
    .EXAMPLE
        .\LogCollector.ps1 -Events System,Application -DestPath C:\Temp\Log
        C:\Temp\Log に System、Application イベント ログをエクスポートします。
        -StartDate および -EndDate が指定されていないため、前日 1 日分が収集対象となります。

    .NOTES
        Windows 10 にて動作確認を実施しています。
        Windows 10 および Windows Server 2016 以降に搭載された PowerShell をサポートします。
        不具合のご連絡は以下の Github リポジトリまでどうぞ。
        https://github.com/yamanyon/LogCollector
        有償サポートのご依頼は m.yamagishi@robinsons.co.jp もしくは twitter: yamanyon までどうぞ。
#>

####################################################################################################
# スクリプト引数定義
####################################################################################################
param (
    # どのイベントログを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は取得しない
    [ValidateSet("Application","Security","System")]
    [array]$Events,

    # いつからのログを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は前日の 0:00:00 を設定
    [DateTime]$StartDate = (Get-Date).Date.AddDays(-1),

    # いつまでのログを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は -StartDate に 1 日加算
    [DateTime]$EndDate = $StartDate.AddDays(1),

    # ログの保存先を指定する
    [parameter(mandatory=$true)]    # null を許容しない
    [String]$DestPath,

    # 保存先ログの保存期間を日数で指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は無限に保存する
    [string]$DestLifeTime
)

####################################################################################################
# スクリプト環境変数定義
####################################################################################################
# PowerShell Cmdlet のエラー発生時の処理を停止にする
$ErrorActionPreference = "Stop"

####################################################################################################
# ログを記録する
# 戻り値: なし
####################################################################################################
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

####################################################################################################
# 指定されたフォルダーパスが正常にアクセスできるか確認し、正常ならフォルダーオブジェクトを返す
# 戻り値: $objFolder 指定されたパスで取得できたフォルダオブジェクト
####################################################################################################
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

####################################################################################################
# 指定されたフォルダーパスを作成する
# 戻り値: $objFolder 指定されたパスで取得できたフォルダオブジェクト
####################################################################################################
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

####################################################################################################
# イベントオブジェクトを取得する
# 戻り値: $objFolder 指定されたパスで取得できたフォルダオブジェクト
####################################################################################################
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

####################################################################################################
# イベントオブジェクトをCSVにエクスポートする
# 戻り値: なし
####################################################################################################
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

####################################################################################################
# 収集開始引数と収集終了日時引数の比較
# 戻り値: なし
####################################################################################################
function funcDateCheck {
    try {
        if ( $StartDate -ge $EndDate ) {
            throw ( "-EndDate ( " + $EndDate.ToString("yyyy/MM/dd HH:mm:ss") + " ) より -StartDate ( " + $StartDate.ToString("yyyy/MM/dd HH:mm:ss") + " ) が古くなくてはいけません。" )
        }
    }
    catch {
        WriteLog "Error" "funcDateCheck error."
    }

    WriteLog "Info" ( $StartDate.ToString("yyyy/MM/dd HH:mm:ss") + " ～ " + $EndDate.ToString("yyyy/MM/dd HH:mm:ss") + " までのログを出力します。" )
}

####################################################################################################
# 収集開始引数と収集終了日時引数の比較
# 戻り値: $strDestinationNowFolder ログの保存先フォルダ文字列
####################################################################################################
function CreateDestFolder {
    WriteLog "Info" "$DestPath が正常に利用できるか確認します。"
    getFolderObject $DestPath | Out-Null

    WriteLog "Info" "$DestPath フォルダーにログを保存するためのフォルダーを作成します。"
    $objDestinationNowFolder = CreateDestinationNowFolder $DestPath
    $strDestinationNowFolder = $objDestinationNowFolder.FullName
    WriteLog "Info" "作成したフォルダー: $strDestinationNowFolder"

    $strDestinationNowFolder
}

####################################################################################################
# イベントログ系の処理
# 戻り値: なし
####################################################################################################
function funcEventLog {
    param (
        $strDestinationNowFolder    # ログを保存するフォルダー
    )

    # イベントログを保存するフォルダーを作成する
    WriteLog "Info" "$strDestinationNowFolder フォルダーにイベント ログを保存するためのフォルダーを作成します。"
    $objWriteFolder = CreateFolder ( $strDestinationNowFolder + "\EventLog" )
    $strWriteFolder = $objWriteFolder.FullName
    WriteLog "Info" "作成したフォルダー: $strWriteFolder"

    WriteLog "Info" "イベントログのエクスポートを開始します。"
    foreach ( $strEventLog in $Events ) {
        GetEventObjects $strEventLog $strWriteFolder
    }
}

####################################################################################################
# ZIP 圧縮処理
# 戻り値: なし
####################################################################################################
function funcCompress {
    param (
        [String]$strSource,     # 圧縮対象フォルダ
        [String]$strDestination # 圧縮ファイル保存先フォルダ
    )

    # $strSource がフォルダーかどうか確認しオブジェクトを生成
    $objSource = getFolderObject $strSource

    # $objSource フォルダーに保存されているファイルオブジェクトを再帰的に取得
    try {
        $objSourceChileItems = Get-ChildItem $objSource
    }
    catch {
        WriteLog "Error" "funcCompress Get-ChildItem $strSource error."
    }

    # 取得できた $objSourceChildItems を String 型配列に放り込む
    $strSourceChildItems = foreach ( $objSourceChildItem in $objSourceChileItems ) {
        $objSourceChildItem.FullName
    }

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

####################################################################################################
# フォルダーの削除
# 戻り値: なし
####################################################################################################
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

####################################################################################################
# 圧縮ファイルの保存期間が切れたら削除する
# 戻り値: なし
####################################################################################################
function DeleteCompressFile {
     # 圧縮ファイルの保存期間が切れたら削除する

}

####################################################################################################
# 主処理
####################################################################################################

WriteLog "Info" "処理を開始します。"

# ふたつの日付引数のチェック
funcDateCheck

# 指定された -DestPath パラメータからフォルダオブジェクトを取得
$strDestinationNowFolder = CreateDestFolder

# イベント ログの出力処理
funcEventLog $strDestinationNowFolder

# この辺りにログ ファイル (所謂テキスト形式ログ) の収集機能も追加したい

# ZIP 圧縮処理
funcCompress $strDestinationNowFolder $DestPath

# 圧縮後のログ削除処理
DeleteFolder $strDestinationNowFolder

# この辺りに保存期間を超過した圧縮ファイルの削除機能も追加したい
DeleteCompressFile $strDestinationNowFolder

WriteLog "Info" "すべての処理が完了しました。"
exit 0

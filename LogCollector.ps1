<#
    .SYNOPSIS
        イベント ログ保存スクリプト "LogCollector.ps1" ver.0.21.20190208 山岸真人@株式会社ロビンソン
     
    .DESCRIPTION
        Windows が出力したイベント ログや、様々なファイルを、指定した期間、指定したフォルダーに保存します。
     
    .PARAMETER Events
        収集するイベント ログの種類を指定します。
        Application、Security、System のいずれかもしくは複数を指定できます。
        省略した場合はイベント ログを収集しません。

    .PARAMETER SourceFolderPath
        収集するファイルをフルパスもしくは相対パスで指定します。
        複数指定する場合はカンマで区切ります。
        スペースを含む場合は "Full Path" もしくは 'Full Path' のように指定します。
        省略した場合はファイルを収集しません。

    .PARAMETER SourceLifeTime
        -SourceFolderPath オプションで指定されたフォルダーに保存されているファイルの保存期間を日数で指定します。
        保存期間は最終更新日を用いて起算されます。
        保存期間を超過したファイルは削除されます。
        保存期間は時分秒を切り捨てて計算します。
        2019/2/12 9:15:00 に -SourceLifeTime 1 を指定して実行した場合、
        2019/2/11 0:00:00 と等しい、もしくはそれより新しい更新日付のファイルは残置され、それより古いファイルは削除されます。
        省略した場合、ゼロもしくはマイナスの値が指定された場合はファイルを削除しません。
    
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
        保存期間は時分秒を切り捨てて計算します。
        2019/2/12 9:15:00 に -ZipLifeTime 1 を指定して実行した場合、
        2019/2/11 0:00:00 と等しい、もしくはそれより新しい更新日付のファイルは残置され、それより古いファイルは削除されます。
        省略した場合、ゼロもしくはマイナスの値が指定された場合は圧縮ファイルを削除しません。

    .INPUTS
        なし。パイプラインからの入力には対応していません。
     
    .OUTPUTS
        正常終了もしくは警告終了した場合は 0 を、異常終了した場合は 1 を返します。
     
    .EXAMPLE
        .\LogCollector.ps1 -Events System,Application -ZipPath C:\Temp\Log
        System および Application イベント ログをエクスポートし、C:\Temp\Log フォルダーに圧縮したものを保存します。
        圧縮ファイルは実行した日時を元に yyyyMMddHHmmss.zip という名称で設定されます。
        -StartDate および -EndDate は指定されていないため、前日 1 日分がエクスポートの対象となります。
        圧縮ファイルには EventLogs フォルダーが生成され、そこに CSV 形式でエクスポートされたイベント ログが保存されます。
        -ZipLifeTime は指定されていないため、C:\Temp\Log フォルダーに保存されたファイルは削除されません。

        .\LogCollector.ps1 -SourceFolderPath "C:\Logs\ABC","C:\Logs\DEF" -SourceLifeTime 7 -ZipPath "D:\ZIP" -ZipLifeTime 365
        C:\Logs\ABC フォルダーおよび C:\Logs\DEF フォルダーに保存されたファイルを C:\ZIP フォルダーに圧縮したものを保存します。
        圧縮ファイルは実行した日時を元に yyyyMMddHHmmss.zip という名称で設定されます。
        -StartDate および -EndDate は指定されていないため、前日 1 日分が圧縮対象となります。
        圧縮ファイルには ABC および DEF フォルダーが生成され、そこに -SourceFolderPath で指定したそれぞれのフォルダーにある圧縮対象のファイルが保存されます。
        -SourceLifeTime に 7 が指定されているため、C:\Logs\ABC および C:\Logs\DEF フォルダーに保存されたファイルは 7 日前の 0:00:00 より古いものが削除されます。
        -ZipLifeTime に 365 が指定されているため、D:\ZIP フォルダーに保存されたファイルは 365 日前の 0:00:00 より古いものが削除されます。

        .\LogCollector.ps1 -ZipPath "D:\ABC" -ZipLifeTime 365
        -ZipPath に D:\ABC が指定されていますが、収集するファイルやイベント ログが指定されていないため、圧縮ファイルは生成されません。
        -ZipLifeTime に 365 が指定されているため、D:\ABC フォルダーに保存されたファイルは 365 日前の 0:00:00 より古いものが削除されます。

    .NOTES
        Windows 10 にて動作確認を実施しています。
        Windows 10 および Windows Server 2016 以降に搭載された PowerShell をサポートします。

        不具合や機能追加のご連絡は以下の Github リポジトリまでどうぞ。
        https://github.com/yamanyon/LogCollector

        有償サポートのご依頼は m.yamagishi@robinsons.co.jp もしくは twitter: yamanyon までどうぞ。
        不具合発生時の問い合わせ窓口の提供等、依頼内容に応じて対応を検討いたします。
#>

<###################################################################################################
v0.20 2018/12/20 初期リリース
v0.21 2019/02/08 ログ収集部分を改良, 主処理系を関数化, コメント改修
v0.22 2019/02/12 -SourceLifeTime パラメーターを実装, エラーメッセージを最適化, ヘルプを最適化
v0.23 xxxx/xx/xx イベント ログの保存フォーマットを CSV 形式と EVTX 形式の両方に対応する
v1.00 xxxx/xx/xx 単体テスト項目作成, 単体テスト実施
###################################################################################################>

<###################################################################################################
スクリプト引数定義
###################################################################################################>
param (
    # どのイベントログを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は取得しない
    [ValidateSet("Application","Security","System")]
    [array]$Events,
    
    # どのフォルダーに保存されたファイルを取得するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は取得しない
    [array]$SourceFolderPath,

    # -SourceFolderPath で指定されたフォルダーに保存されているファイルをそのフォルダで何日間保存するかを指定する
    [parameter(mandatory=$false)]   # null を許容する、null の場合は削除しない
    [int]$SourceLifeTime,

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
    [parameter(mandatory=$false)]   # null を許容する、null の場合は削除しない
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
    
    # 保存期間を超過したZIPファイルを削除する
    DeleteExpiredFile $ZipPath $ZipLifeTime

    # 保存期間を超過したファイルを削除する
    foreach ( $strPath in $SourceFolderPath ) {
        DeleteExpiredFile $strPath $SourceLifeTime
    }
    
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
    $format = "yyyy/MM/dd HH:mm:ss"

    try {
        if ( $StartDate -ge $EndDate ) {
            throw ( "-EndDate ( " + $EndDate.ToString($format) + " ) より -StartDate ( " + $StartDate.ToString($format) + " ) が新しいため、処理できません。" )
        }
    }
    catch {
        WriteLog "Error" "funcDateCheck error."
    }

    WriteLog "Info" ( $StartDate.ToString($format) + " ～ " + $EndDate.ToString($format) + " までのイベント ログやファイルを収集します。" )
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
        WriteLog "Error" "指定されたパス $strFolder は利用できません。"
    }

    # フォルダー文字列が利用できるなら、アイテムを取得し、アイテムがフォルダーであることを確認する
    if ( $bTestResult ) {
        # 指定された文字列で Get-Item してみる
        try {
            $objFolder = Get-Item $strFolder
        }
        catch {
            WriteLog "Error" "指定されたパス $strFolder へアクセスできません。"
        }

        # Get-Item できたらそれがコンテナオブジェクトか確認する
        try {
            if ( !$objFolder.PSIsContainer ) {
                throw
            }
        }
        catch {
            WriteLog "Error" "指定されたパス $strFolder はフォルダーではありません。"
        }
    }
    # フォルダー文字列が利用できないのなら、フォルダーを試しに作成してみる
    else {
        $objFolder = CreateFolder $strFolder
    }

    # エラー無く処理できたら、$objFolder オブジェクトを返す
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
        WriteLog "Error" "New-Item Cmdlet を利用し、指定されたフォルダー $strFolder を作成しようとして、問題が発生しました。"
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
        WriteLog "Warn" "イベント ログ $strEventLog を参照しようとして失敗しました。イベント ログの保存処理をスキップします。"
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
    WriteLog "Info" "フォルダー $strPath にイベント ログ $strEventLog をエクスポートします。"
    try {
        $objEvents | Export-Csv -Path $strPath -Encoding Default -NoTypeInformation
    }
    catch {
        WriteLog "Error" "フォルダー $strPath にイベント ログ $strEvdentLog をエクスポートしようとして失敗しました。"
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
            throw
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
        WriteLog "Warn" "テンポラリとして作成したフォルダー $strSource へアクセスできません。圧縮処理をスキップします。"
        return
    }

    # $objSource フォルダーに保存されているオブジェクトにファイルがあることを確認
    try {
        if ( $null -eq $objSourceFiles ) {
            throw
        }
    }
    catch {
        WriteLog "Warn" "テンポラリとして作成したフォルダー $strSource にファイルがありません。圧縮処理をスキップします。"
        return
    }
        
    # $objSource フォルダーに保存されているオブジェクトを取得
    try {
        $objSourceChileItems = Get-ChildItem $objSource
    }
    catch {
        WriteLog "Error" "テンポラリとして作成したフォルダー $strSource からアイテムを収集しようとして失敗しました。圧縮処理をスキップします。"
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
        WriteLog "Error" "テンポラリとして作成したフォルダー $strSource のアイテムを $strDestFileFullPath に圧縮しようとして失敗しました。"
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
        WriteLog "Error" "指定されたフォルダー $strFolderPath を削除できませんでした。"
    }
}

<###################################################################################################
ファイルの保存期間を DateTime 型に変換する
戻り値: $dtLifeTime ファイルの保存期間
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
        WriteLog "Error" "有効期限の切れたファイルを $strFolder から収集しようとして失敗しました。"
    }

    try {
        $objItems | ForEach-Object { Remove-Item $_.FullName -Force }
    }
    catch {
        WriteLog "Error" "有効期限の切れたファイルを $strFolder から削除しようとして失敗しました。"
    }
}

<###################################################################################################
主処理実行
###################################################################################################>
LogCollector

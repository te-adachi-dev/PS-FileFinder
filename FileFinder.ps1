# ▲ここから保存してください▲

function Maximize-Window {
    $sig = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    Add-Type -Name Win -MemberDefinition $sig -Namespace Native
    [Native.Win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 3) | Out-Null
}

Maximize-Window

$searchPattern = Read-Host "検索したいファイル名またはパターンを入力してください (例: 'test' や '.*\.txt$' など)"
if ([string]::IsNullOrWhiteSpace($searchPattern)) {
    Write-Host "検索パターンが空でした。スクリプトを終了します。" -ForegroundColor Red
    return
}

$results = @()
$drives = Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object DeviceID, DriveType

foreach ($driveInfo in $drives) {
    $drive = $driveInfo.DeviceID
    $driveType = $driveInfo.DriveType

    if ($driveType -eq 5) {
        Write-Host "ドライブ $drive はCD/DVDドライブのためスキップします。" -ForegroundColor Yellow
        continue
    }

    Write-Host "検索中: $drive" -ForegroundColor Cyan

    try {
        $null = Get-Item -Path $drive -ErrorAction Stop
    }
    catch {
        Write-Host "$drive は準備できていないか、アクセスできません。スキップします。" -ForegroundColor Red
        continue
    }

    try {
        $directories = Get-ChildItem -Path $drive -Directory -Recurse -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "$drive のディレクトリ取得でエラーが発生しました。スキップします。" -ForegroundColor Red
        continue
    }

    $totalDirs = $directories.Count
    $currentDirIndex = 0

    foreach ($directory in $directories) {
        $currentDirIndex++
        Write-Progress -Activity "検索中 (Drive: $drive)" `
                       -Status "$($directory.FullName)" `
                       -PercentComplete ([math]::Floor($currentDirIndex / $totalDirs * 100))

        try {
            $files = Get-ChildItem -Path $directory.FullName -File -ErrorAction SilentlyContinue
        }
        catch {
            continue
        }

        $pattern = $searchPattern
        $matchedFiles = $files | Where-Object { $_.Name -match $pattern }

        if ($matchedFiles) {
            $results += $matchedFiles | Select-Object FullName
        }
    }
}

Write-Progress -Activity "完了" -Status "検索処理が終了しました。" -Completed

if ($results) {
    Write-Host "`n検索結果:" -ForegroundColor Green
    $results | ForEach-Object { Write-Host $_.FullName }

    $results | Select-Object -ExpandProperty FullName | Set-Clipboard
    Write-Host "`n結果がクリップボードにコピーされました。" -ForegroundColor Cyan
} else {
    Write-Host "`n指定されたパターンに一致するファイルは見つかりませんでした。" -ForegroundColor Red
}

Write-Host "`nEnterキーを押して終了してください..."
Read-Host | Out-Null

# ▲ここまで保存してください▲

# PowerShell 7以降の場合、ForEach-Object -Parallelを使用可能
# PowerShell 5.1の場合、ThreadJobまたはRunspacePoolを使用

# ウィンドウを最大化する関数
function Maximize-Window {
    $sig = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    Add-Type -Name Win -MemberDefinition $sig -Namespace Native
    [Native.Win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 3) | Out-Null
}

# ログファイルの設定
$logPath = "${env:TEMP}\FileSearch_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$global:syncHash = [hashtable]::Synchronized(@{})

# ログ出力関数
function Write-SearchLog {
    param(
        [string]$Message,
        [string]$Type = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Type] $Message"
    Add-Content -Path $logPath -Value $logMessage -Force
    
    # コンソールにも表示
    switch ($Type) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        default { Write-Host $logMessage -ForegroundColor Cyan }
    }
}

# ウィンドウを最大化
Maximize-Window

# 検索パターンの入力を求める
$searchPattern = Read-Host "検索したいファイル名またはパターンを入力してください (例: 'test' や '.*\.txt$' など)"
if ([string]::IsNullOrWhiteSpace($searchPattern)) {
    Write-SearchLog "検索パターンが空でした。スクリプトを終了します。" -Type "ERROR"
    return
}

Write-SearchLog "検索開始: パターン '$searchPattern'"

# システム上の全てのドライブを取得
$drives = Get-CimInstance -ClassName Win32_LogicalDisk | 
    Where-Object { $_.DriveType -in @(2, 3, 4) } | # Removable, Local, Network
    Select-Object DeviceID, DriveType, Size, FreeSpace

Write-SearchLog "検出されたドライブ数: $($drives.Count)"

# PowerShellバージョンチェック
$psVersion = $PSVersionTable.PSVersion.Major
Write-SearchLog "PowerShell Version: $psVersion"

if ($psVersion -ge 7) {
    # PowerShell 7以降の場合、ForEach-Object -Parallelを使用
    Write-SearchLog "PowerShell 7+検出: ForEach-Object -Parallelを使用します"
    
    # 進捗追跡用の同期ハッシュテーブル
    $progressHash = [hashtable]::Synchronized(@{})
    $resultsHash = [hashtable]::Synchronized(@{})
    
    # ドライブごとの処理を並列実行
    $jobs = $drives | ForEach-Object -ThrottleLimit 5 -AsJob -Parallel {
        $drive = $_.DeviceID
        $searchPattern = $using:searchPattern
        $syncProgress = $using:progressHash
        $syncResults = $using:resultsHash
        $logPath = $using:logPath
        
        # 進捗情報の初期化
        $syncProgress[$drive] = @{
            Status = "開始"
            Progress = 0
            FileCount = 0
            CurrentPath = ""
            StartTime = Get-Date
        }
        
        # ドライブ固有の結果配列を初期化
        $syncResults[$drive] = [System.Collections.ArrayList]::new()
        
        try {
            # ドライブアクセス確認
            $null = Get-Item -Path $drive -ErrorAction Stop
            
            # ファイル検索の実行
            $allFiles = Get-ChildItem -Path $drive -File -Recurse -ErrorAction SilentlyContinue
            $totalFiles = $allFiles.Count
            $currentIndex = 0
            $foundCount = 0
            
            foreach ($file in $allFiles) {
                $currentIndex++
                
                # 進捗更新（100ファイルごと）
                if ($currentIndex % 100 -eq 0) {
                    $syncProgress[$drive] = @{
                        Status = "検索中"
                        Progress = [math]::Round(($currentIndex / $totalFiles) * 100, 2)
                        FileCount = $foundCount
                        CurrentPath = $file.DirectoryName
                        StartTime = $syncProgress[$drive].StartTime
                    }
                }
                
                # パターンマッチング
                if ($file.Name -match $searchPattern) {
                    $null = $syncResults[$drive].Add($file.FullName)
                    $foundCount++
                }
            }
            
            # 完了状態を更新
            $syncProgress[$drive] = @{
                Status = "完了"
                Progress = 100
                FileCount = $foundCount
                CurrentPath = ""
                StartTime = $syncProgress[$drive].StartTime
                EndTime = Get-Date
            }
            
            Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [SUCCESS] ドライブ ${drive}: 検索完了 (${foundCount}件)"
            
        } catch {
            $syncProgress[$drive] = @{
                Status = "エラー"
                Progress = 0
                FileCount = 0
                CurrentPath = $_.Exception.Message
                StartTime = $syncProgress[$drive].StartTime
                EndTime = Get-Date
            }
            Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR] ドライブ ${drive}: $_"
        }
    }
    
    # 進捗モニタリング
    Write-Host "`n並列検索実行中..." -ForegroundColor Green
    while ($jobs.State -eq 'Running') {
        Clear-Host
        Write-Host "=== ファイル検索進捗状況 ===" -ForegroundColor Cyan
        Write-Host "検索パターン: $searchPattern" -ForegroundColor Yellow
        Write-Host "ログファイル: $logPath" -ForegroundColor Gray
        Write-Host ("-" * 60)
        
        foreach ($drive in $drives.DeviceID) {
            if ($progressHash.ContainsKey($drive)) {
                $info = $progressHash[$drive]
                $status = $info.Status
                $progress = $info.Progress
                $fileCount = $info.FileCount
                $currentPath = $info.CurrentPath
                
                Write-Host "`nドライブ ${drive}:" -ForegroundColor White
                Write-Host "  状態: $status" -ForegroundColor $(
                    switch ($status) {
                        "完了" { "Green" }
                        "エラー" { "Red" }
                        default { "Yellow" }
                    }
                )
                Write-Host "  進捗: ${progress}%"
                Write-Host "  検出ファイル数: $fileCount"
                if ($currentPath) {
                    Write-Host "  現在の検索パス: $currentPath" -ForegroundColor Gray
                }
            }
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    # ジョブの完了を待つ
    $null = $jobs | Wait-Job
    $jobs | Remove-Job
    
} else {
    # PowerShell 5.1の場合、RunspacePoolを使用
    Write-SearchLog "PowerShell 5.1検出: RunspacePoolを使用します"
    
    # RunspacePoolの作成
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min($drives.Count, 5))
    $runspacePool.Open()
    
    # 各ドライブの検索を実行
    $runspaces = @()
    
    $scriptBlock = {
        param($drive, $searchPattern, $logPath)
        
        $results = @()
        $status = @{
            Drive = $drive
            Status = "開始"
            FileCount = 0
            StartTime = Get-Date
        }
        
        try {
            # ドライブアクセス確認
            $null = Get-Item -Path $drive -ErrorAction Stop
            
            # ファイル検索
            Get-ChildItem -Path $drive -File -Recurse -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -match $searchPattern } | 
                ForEach-Object {
                    $results += $_.FullName
                }
            
            $status.Status = "完了"
            $status.FileCount = $results.Count
            $status.EndTime = Get-Date
            
            Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [SUCCESS] ドライブ ${drive}: 検索完了 ($($results.Count)件)"
            
        } catch {
            $status.Status = "エラー"
            $status.Error = $_.Exception.Message
            Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR] ドライブ ${drive}: $_"
        }
        
        return @{
            Drive = $drive
            Results = $results
            Status = $status
        }
    }
    
    # 各ドライブに対してRunspaceを作成
    foreach ($driveInfo in $drives) {
        $drive = $driveInfo.DeviceID
        
        $powershell = [powershell]::Create()
        $powershell.RunspacePool = $runspacePool
        
        [void]$powershell.AddScript($scriptBlock)
        [void]$powershell.AddArgument($drive)
        [void]$powershell.AddArgument($searchPattern)
        [void]$powershell.AddArgument($logPath)
        
        $runspaces += [PSCustomObject]@{
            PowerShell = $powershell
            Handle = $powershell.BeginInvoke()
            Drive = $drive
        }
        
        Write-SearchLog "ドライブ $drive の検索を開始しました"
    }
    
    # 完了待機と進捗表示
    Write-Host "`n並列検索実行中..." -ForegroundColor Green
    $completed = @()
    
    while ($runspaces.Count -gt $completed.Count) {
        foreach ($runspace in $runspaces) {
            if ($runspace.Handle.IsCompleted -and $runspace.Drive -notin $completed) {
                $completed += $runspace.Drive
                Write-Host "ドライブ $($runspace.Drive) の検索が完了しました" -ForegroundColor Green
            }
        }
        
        Write-Host "`r進捗: $($completed.Count)/$($runspaces.Count) ドライブ完了" -NoNewline
        Start-Sleep -Milliseconds 500
    }
    
    Write-Host "`n"
    
    # 結果の収集
    $resultsHash = @{}
    foreach ($runspace in $runspaces) {
        try {
            $result = $runspace.PowerShell.EndInvoke($runspace.Handle)
            if ($result) {
                $resultsHash[$result.Drive] = $result.Results
            }
        } catch {
            Write-SearchLog "ドライブ $($runspace.Drive) の結果取得でエラー: $_" -Type "ERROR"
        } finally {
            $runspace.PowerShell.Dispose()
        }
    }
    
    $runspacePool.Close()
    $runspacePool.Dispose()
}

# 結果の統合と表示
Write-Host "`n=== 検索結果 ===" -ForegroundColor Green
$allResults = @()

foreach ($drive in $drives.DeviceID) {
    if ($resultsHash.ContainsKey($drive) -and $resultsHash[$drive].Count -gt 0) {
        Write-Host "`nドライブ ${drive}: $($resultsHash[$drive].Count)件" -ForegroundColor Cyan
        foreach ($file in $resultsHash[$drive]) {
            Write-Host "  $file"
            $allResults += $file
        }
    }
}

if ($allResults.Count -gt 0) {
    # 結果をクリップボードにコピー
    $allResults | Set-Clipboard
    Write-SearchLog "総検出ファイル数: $($allResults.Count)" -Type "SUCCESS"
    Write-Host "`n結果がクリップボードにコピーされました。" -ForegroundColor Cyan
} else {
    Write-SearchLog "指定されたパターンに一致するファイルは見つかりませんでした。" -Type "WARNING"
}

Write-SearchLog "検索処理完了"
Write-Host "`nログファイル: $logPath" -ForegroundColor Gray
Write-Host "`nEnterキーを押して終了してください..."
try {
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} catch {
    # ISEやVS Codeターミナルの場合は代替手段を使用
    Read-Host "Enterキーを押してください"
}

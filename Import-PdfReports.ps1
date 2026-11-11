#requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$InputPath = @(),
    [ValidateRange(1,2147483647)][int]$StartPage = 2,
    [string]$DataRoot = '',
    [string]$OutputRoot = ''
)
$ErrorActionPreference = 'Stop'
$lock=$null; $manifestPath=$null
$committed=$false; $exitCode=0
$timer=[Diagnostics.Stopwatch]::StartNew()
$utf8=[Text.UTF8Encoding]::new($false)
$batchId=(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8)
$importedAt=[DateTimeOffset]::Now.ToString('o')
$manifest=[ordered]@{BatchId=$batchId;ImportedAt=$importedAt;Status='Processing';WorkbookCommitted=$false;StartPage=$StartPage;Current=$null;Files=@();Error=''}
function Save-ReportCheckpoint([object[]]$Results,[object[]]$Logs) {
    $checkpointPackage=$null; $checkpointPath=$null
    try {
        if([IO.File]::Exists($workbook)) { $checkpointPackage=[OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]::new($workbook)) }
        else { $checkpointPackage=[OfficeOpenXml.ExcelPackage]::new() }
        Add-ReportResults $checkpointPackage $Results $Logs $importedAt
        $checkpointPath=Join-Path $DataRoot ('pending-'+[guid]::NewGuid().ToString('N')+'.xlsx')
        $checkpointPackage.SaveAs([IO.FileInfo]::new($checkpointPath))
        $checkpointPackage.Dispose(); $checkpointPackage=$null
        if([IO.File]::Exists($workbook)) { [IO.File]::Replace($checkpointPath,$workbook,(Join-Path $DataRoot 'Reports.previous.xlsx')) }
        else { [IO.File]::Move($checkpointPath,$workbook) }
        $checkpointPath=$null
    } finally {
        if($checkpointPackage) { $checkpointPackage.Dispose() }
        if($checkpointPath -and [IO.File]::Exists($checkpointPath)) { [IO.File]::Delete($checkpointPath) }
    }
}
try {
    Write-Host '啟動 PDF 表格萃取工具，正在載入離線函式庫…'
    if (!$InputPath.Count) { $InputPath=@((Join-Path $PSScriptRoot 'Input')) }
    if (!$DataRoot) { $DataRoot=Join-Path $PSScriptRoot 'Data' }
    Import-Module -Name (Join-Path $PSScriptRoot 'Extractor.Core.psm1') -Force -ErrorAction Stop
    Initialize-Extractor
    $files=[Collections.Generic.List[string]]::new()
    foreach($inputItem in $InputPath) {
        $item=Get-Item -LiteralPath $inputItem -ErrorAction Stop
        if($item.PSIsContainer) {
            foreach($file in Get-ChildItem -LiteralPath $item.FullName -File -Filter '*.pdf' | Sort-Object FullName) { $files.Add($file.FullName) }
        } else {
            if($item.Extension -ine '.pdf') { throw "不是 PDF 檔案：$($item.FullName)" }
            $files.Add($item.FullName)
        }
    }
    if(!$files.Count) { throw '找不到 PDF。請將檔案放入 Input，或使用 -InputPath 指定檔案／資料夾。' }
    $DataRoot=[IO.Path]::GetFullPath($DataRoot)
    if(!$OutputRoot) { $OutputRoot=Join-Path $DataRoot 'Output' }
    $OutputRoot=[IO.Path]::GetFullPath($OutputRoot)
    [IO.Directory]::CreateDirectory($DataRoot) | Out-Null
    try { $lock=[IO.File]::Open((Join-Path $DataRoot '.session.lock'),'OpenOrCreate','ReadWrite','None') }
    catch { throw '資料夾正在被另一個匯入程序使用，請等候完成後再執行。' }
    $workbook=Join-Path $DataRoot 'Reports.xlsx'
    if([IO.File]::Exists($workbook)) {
        try { $check=[IO.File]::Open($workbook,'Open','ReadWrite','None'); $check.Dispose() }
        catch { throw "歷史 Excel 已被鎖定或無法寫入，請先關閉 Excel：$workbook" }
    }
    $output=Join-Path $OutputRoot $batchId
    [IO.Directory]::CreateDirectory($output) | Out-Null
    $manifestPath=Join-Path $output 'manifest.json'
    [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8)
    Write-Host ("找到 {0} 份 PDF；從實際第 {1} 頁開始。" -f $files.Count,$StartPage)
    Write-Host ("資料位置：{0}" -f $DataRoot)
    Write-Host ("本次進度紀錄：{0}" -f $manifestPath)
    Write-Host '每份 PDF 完整解析成功後立即儲存 Excel。'
    $logs=[Collections.Generic.List[object]]::new()
    $pendingLogs=[Collections.Generic.List[object]]::new()
    $fileIndex=0
    foreach($file in $files) {
        $fileIndex++
        $result=$null
        $log=[pscustomobject]@{BatchId=$batchId;ImportedAt=$importedAt;SourceFile=[IO.Path]::GetFileName($file);SourcePath=$file;SHA256='';StartPage=$StartPage;Status='Failed';Rows=0;Error=''}
        try {
            $manifest.Current=[ordered]@{File=$file;FileIndex=$fileIndex;FileCount=$files.Count;Page=0;TotalPages=0;ParsedRows=0;Phase='OpeningPdf';ElapsedSeconds=[Math]::Round($timer.Elapsed.TotalSeconds,1);UpdatedAt=[DateTimeOffset]::Now.ToString('o')}
            [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8)
            Write-Host ("[{0}/{1}] 正在開啟：{2}" -f $fileIndex,$files.Count,$log.SourceFile)
            $onProgress={
                param($page,$totalPages,$rowCount,$phase)
                $manifest.Current.Page=$page; $manifest.Current.TotalPages=$totalPages; $manifest.Current.ParsedRows=$rowCount; $manifest.Current.Phase=$phase
                $manifest.Current.ElapsedSeconds=[Math]::Round($timer.Elapsed.TotalSeconds,1); $manifest.Current.UpdatedAt=[DateTimeOffset]::Now.ToString('o')
                [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8)
                $finished=$page-$StartPage+$(if($phase-eq 'PageComplete'){1}else{0})
                $percent=[Math]::Min(100,[int](100*$finished/($totalPages-$StartPage+1)))
                Write-Progress -Id 1 -Activity ("[{0}/{1}] {2}" -f $fileIndex,$files.Count,$log.SourceFile) -Status ("第 {0}/{1} 頁，已解析 {2} 筆，經過 {3} 秒" -f $page,$totalPages,$rowCount,$manifest.Current.ElapsedSeconds) -PercentComplete $percent
                if ($phase-eq 'PageComplete' -and ($page-eq $StartPage -or $page%25-eq 0 -or $page-eq $totalPages)) {
                    Write-Host ("  第 {0}/{1} 頁完成，已解析 {2} 筆（累計 {3} 秒）。" -f $page,$totalPages,$rowCount,$manifest.Current.ElapsedSeconds)
                }
            }
            $result=Read-ReportFile $file $StartPage $batchId $importedAt $onProgress
        } catch {
            $log.Error=$_.Exception.Message
            Write-Warning ("未匯入：{0}；{1}" -f $log.SourceFile,$log.Error)
            $exitCode=1
        }
        $logs.Add($log)
        $pendingLogs.Add($log)
        if($result) {
            $log.SHA256=$result.Hash; $log.Rows=$result.Records.Count; $log.Status='PendingSave'
            $manifest.Current.Phase='SavingWorkbook'
            $manifest.Files=@($logs.ToArray())
            [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8)
            Write-Host ("解析完成：{0}，{1} 筆，正在儲存 Excel…" -f $log.SourceFile,$log.Rows)
            try {
                # The success log and this PDF's complete records are committed together.
                $checkpointLogs=@($pendingLogs.ToArray())
                $savedLog=$log.PSObject.Copy(); $savedLog.Status='Success'
                $checkpointLogs[$checkpointLogs.Count-1]=$savedLog
                Save-ReportCheckpoint @($result) $checkpointLogs
            } catch {
                $log.Status='Failed'; $log.Rows=0; $log.Error='Excel 儲存失敗：'+$_.Exception.Message
                $manifest.Files=@($logs.ToArray())
                throw
            }
            $log.Status='Success'
            $committed=$true; $manifest.WorkbookCommitted=$true
            $pendingLogs.Clear()
            $result=$null
            Write-Host ("已儲存：{0}，{1} 筆；可繼續處理下一份。" -f $log.SourceFile,$log.Rows)
        }
        Write-Progress -Id 1 -Activity $log.SourceFile -Completed
        $manifest.Current.Phase='FileComplete'
        $manifest.Current.ElapsedSeconds=[Math]::Round($timer.Elapsed.TotalSeconds,1); $manifest.Current.UpdatedAt=[DateTimeOffset]::Now.ToString('o')
        $manifest.Files=@($logs.ToArray())
        [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8)
    }
    if($committed) {
        if($pendingLogs.Count) {
            $manifest.Current.Phase='SavingWorkbook'
            [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8)
            Save-ReportCheckpoint @() $pendingLogs.ToArray()
            $pendingLogs.Clear()
        }
        $manifest.Status=if($exitCode){'PartialSuccess'}else{'Success'}
    } else { $manifest.Status='Failed'; $manifest.Error='所有 PDF 均解析失敗，未變更歷史 Excel。' }
    if ($manifest.Current) { $manifest.Current.Phase='Finished'; $manifest.Current.ElapsedSeconds=[Math]::Round($timer.Elapsed.TotalSeconds,1); $manifest.Current.UpdatedAt=[DateTimeOffset]::Now.ToString('o') }
    [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8)
    [pscustomobject]@{
        BatchId=$batchId; Workbook=$workbook; WorkbookCommitted=$committed; Output=$output
        SuccessfulFiles=@($logs | Where-Object Status -eq 'Success').Count
        FailedFiles=@($logs | Where-Object Status -eq 'Failed').Count
        AddedRows=($logs | Measure-Object Rows -Sum).Sum
        Status=$manifest.Status
    }
} catch {
    $exitCode=2
    $manifest.Status='Failed'; $manifest.WorkbookCommitted=$committed; $manifest.Error=$_.Exception.Message
    if($manifestPath) {
        try { [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8) } catch {}
    }
    Write-Error -Message ("匯入失敗：{0}。歷史 Excel 已更新：{1}" -f $manifest.Error,$committed) -ErrorAction Continue
} finally {
    if ($manifestPath -and $manifest.Status -eq 'Processing') {
        $manifest.Status='Interrupted'; $manifest.Error='執行未完成；請確認 WorkbookCommitted 狀態及最後進度。'
        try { [IO.File]::WriteAllText($manifestPath,(ConvertTo-Json -InputObject $manifest -Depth 8),$utf8) } catch {}
    }
    Write-Progress -Id 1 -Activity 'PDF 表格萃取' -Completed
    if($lock) { $lock.Dispose() }
}
exit $exitCode

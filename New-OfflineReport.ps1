#requires -Version 5.1
[CmdletBinding()]
param([string]$WorkbookPath='', [string]$OutputPath='', [string]$Title='事件分析報表')
$ErrorActionPreference='Stop'
$package=$null; $pending=$null; $stream=$null; $inputStream=$null
# First entry is the grouping key; remaining entries define the detail order.
$displayFormats=@{
    'Active Alerts (Trouble)'=@('Source','Time Generated','Alert Message')
    'All First Time Resource Access'=@('Reference Resource','Resource Accessed Time','Message')
    'All Threshold Breach Activity'=@('Name','Time Generated','Mean Count','Threshold Count','Message')
    'All Unusual Time Activity'=@('Name','Time Generated','General Start Time','General End Time','Message')
    'First Time Remote Access on Host'=@('Unusual Remote Access From Computer','Resource Accessed Time','Message')
    'Unusual Logon Activity Time'=@('User Name','Time Generated','General Start Time','General End Time','Message')
    'Unusual Logon Activity Time on Host'=@('Computer Name','Time Generated','General Start Time','General End Time','Message')
    'Unusual User Activity Time'=@('User Name','Time Generated','General Start Time','General End Time','Message')
    'Unusual Volume of Lockout'=@('User Name','Time Generated','Mean Count','Threshold Count','Message')
    'Unusual Volume of Logon Failure'=@('User Name','Time Generated','Mean Count','Threshold Count','Message')
    'Unusual Volume of Logon Failure on Host'=@('Computer Name','Time Generated','Mean Count','Threshold Count','Message')
    'Unusual Volume of User Management Activity'=@('User Name','Time Generated','Mean Count','Threshold Count','Message')
}
try {
    if (!$WorkbookPath) { $WorkbookPath=Join-Path $PSScriptRoot 'Data/Reports.xlsx' }
    if (!$OutputPath) { $OutputPath=Join-Path $PSScriptRoot 'Data/Report.html' }
    $WorkbookPath=[IO.Path]::GetFullPath($WorkbookPath)
    $OutputPath=[IO.Path]::GetFullPath($OutputPath)
    if ($OutputPath -eq $WorkbookPath -or [IO.Path]::GetExtension($OutputPath) -ine '.html') { throw '輸出位置必須是獨立的 .html 檔案。' }
    Write-Host '正在讀取 Excel，建立離線報表…'
    Import-Module (Join-Path $PSScriptRoot 'Extractor.Core.psm1') -Force
    Initialize-Extractor
    # Read a snapshot; do not modify the historical workbook.
    $stream=[IO.MemoryStream]::new()
    $inputStream=[IO.File]::Open($WorkbookPath,'Open','Read','ReadWrite')
    $inputStream.CopyTo($stream); $inputStream.Dispose(); $inputStream=$null
    $stream.Position=0
    try { $package=[OfficeOpenXml.ExcelPackage]::new($stream) } catch { $stream.Dispose(); throw }
    $index=$package.Workbook.Worksheets['_Reports']
    if (!$index -or !$index.Dimension) { throw 'Excel 缺少 _Reports 索引，請使用本工具建立的 Reports.xlsx。' }
    $reports=[Collections.Generic.List[object]]::new()
    $total=0
    for($ir=2;$ir -le $index.Dimension.End.Row;$ir++) {
        $reportTitle=[string]$index.Cells[$ir,1].Value
        $sheetName=[string]$index.Cells[$ir,2].Value
        $sheet=$package.Workbook.Worksheets[$sheetName]
        if (!$reportTitle -or !$sheet -or !$sheet.Dimension) { throw "報告索引不完整：$sheetName" }
        $columns=[Collections.Generic.List[string]]::new()
        $positions=[Collections.Generic.List[int]]::new()
        $sourceFile=0; $sourcePage=0
        for($c=1;$c -le $sheet.Dimension.End.Column;$c++) {
            $name=[string]$sheet.Cells[1,$c].Value
            if ($name -ceq '__SourceFile') { $sourceFile=$c }
            if ($name -ceq '__SourcePage') { $sourcePage=$c }
            if ($name.StartsWith('__')) { continue }
            if (!$name -or $columns.Contains($name)) { throw "工作表 $sheetName 存在空白或重複表頭。" }
            $columns.Add($name); $positions.Add($c)
        }
        if (!$columns.Count) { throw "工作表 $sheetName 沒有 PDF 欄位。" }
        $columnLookup=@{}
        for($c=0;$c -lt $columns.Count;$c++) {
            $normalized=([regex]::Replace($columns[$c],'\s+',' ')).Trim()
            if($columnLookup.ContainsKey($normalized)) { throw "報告 $reportTitle 存在重複欄位：$normalized" }
            $columnLookup[$normalized]=$c
        }
        $keyIndex=0
        $detailIndices=[int[]]@(for($c=1;$c -lt $columns.Count;$c++) { $c })
        if($displayFormats.ContainsKey($reportTitle)) {
            $format=$displayFormats[$reportTitle]
            $missing=@($format | Where-Object { !$columnLookup.ContainsKey($_) })
            if($missing.Count) { throw "報告 $reportTitle 缺少必要欄位：$($missing -join '、')" }
            $keyIndex=[int]$columnLookup[$format[0]]
            $detailIndices=[int[]]@($format | Select-Object -Skip 1 | ForEach-Object { $columnLookup[$_] })
        }
        $timeColumns=@(for($c=0;$c -lt $columns.Count;$c++) {
            $normalized=([regex]::Replace($columns[$c],'\s+',' ')).Trim()
            if ($normalized -in @('Time Generated','Resource Accessed Time')) { $c }
        })
        if ($timeColumns.Count -ne 1) { throw "工作表 $sheetName 必須有且只有一個 Time Generated 或 Resource Accessed Time 欄位。" }
        $rows=[Collections.Generic.List[object]]::new()
        for($r=2;$r -le $sheet.Dimension.End.Row;$r++) {
            $values=[string[]]::new($columns.Count)
            for($c=0;$c -lt $columns.Count;$c++) { $values[$c]=[string]$sheet.Cells[$r,$positions[$c]].Value }
            $rows.Add([pscustomobject]@{values=$values;file=$(if($sourceFile){[string]$sheet.Cells[$r,$sourceFile].Value}else{''});page=$(if($sourcePage){[string]$sheet.Cells[$r,$sourcePage].Value}else{''})})
        }
        $reports.Add([pscustomobject]@{id='report-'+($ir-1);title=$reportTitle;sheet=$sheetName;columns=$columns.ToArray();keyIndex=$keyIndex;detailIndices=$detailIndices;timeIndex=[int]$timeColumns[0];rows=$rows.ToArray()})
        $total+=$rows.Count
        Write-Host ("  {0}：{1} 筆，識別欄位 {2}" -f $reportTitle,$rows.Count,$columns[$keyIndex])
    }
    $data=[ordered]@{title=$Title;generatedAt=[DateTimeOffset]::Now.ToString('o');workbook=[IO.Path]::GetFileName($WorkbookPath);reports=$reports.ToArray()}
    $json=ConvertTo-Json -InputObject $data -Depth 12 -Compress
    # Data must never terminate the inert script element, even if a PDF contains HTML.
    $json=$json.Replace('&','\u0026').Replace('<','\u003c').Replace('>','\u003e').Replace([string][char]0x2028,'\u2028').Replace([string][char]0x2029,'\u2029')
    $template=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'report/index.html'))
    $css=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'report/report.css'))
    $js=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'report/report.js'))
    $html=$template.Replace('<!--REPORT_TITLE-->',[Net.WebUtility]::HtmlEncode($Title)).Replace('/*REPORT_CSS*/',$css).Replace('/*REPORT_JS*/',$js).Replace('<!--REPORT_DATA-->',$json)
    $directory=[IO.Path]::GetDirectoryName($OutputPath)
    [IO.Directory]::CreateDirectory($directory)|Out-Null
    $pending=Join-Path $directory ('report-pending-'+[guid]::NewGuid().ToString('N')+'.html')
    [IO.File]::WriteAllText($pending,$html,[Text.UTF8Encoding]::new($true))
    # Windows PowerShell 5.1 converts ordinary $null to an empty string for this
    # string parameter. Pass a real null so .NET does not treat it as a backup path.
    if ([IO.File]::Exists($OutputPath)) { [IO.File]::Replace($pending,$OutputPath,[System.Management.Automation.Language.NullString]::Value) }
    else { [IO.File]::Move($pending,$OutputPath) }
    $pending=$null
    Write-Output ([pscustomobject]@{Report=$OutputPath;Worksheets=$reports.Count;Records=$total})
} finally {
    if($package) { $package.Dispose() }
    if($inputStream) { $inputStream.Dispose() }
    if($stream) { $stream.Dispose() }
    if($pending -and [IO.File]::Exists($pending)) { [IO.File]::Delete($pending) }
}

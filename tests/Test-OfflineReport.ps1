#requires -Version 5.1
[CmdletBinding()]
param([string]$Browser='C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',[string]$ActualReport='')
$ErrorActionPreference='Stop'
$tool=Split-Path $PSScriptRoot -Parent
$root=Join-Path ([IO.Path]::GetTempPath()) ('pdf-report-ui-'+[guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($root)|Out-Null
Import-Module (Join-Path $tool 'Extractor.Core.psm1') -Force
Initialize-Extractor
$workbook=Join-Path $root 'fixture.xlsx';$html=Join-Path $root 'fixture.html'
$p=[OfficeOpenXml.ExcelPackage]::new()
try {
    $index=$p.Workbook.Worksheets.Add('_Reports')
    $index.Cells[1,1].Value='ReportTitle';$index.Cells[1,2].Value='SheetName'
    $index.Cells[2,1].Value='Logon fixture';$index.Cells[2,2].Value='Logons'
    $index.Cells[3,1].Value='Resource fixture';$index.Cells[3,2].Value='Resources'
    $s=$p.Workbook.Worksheets.Add('Logons')
    $headers=@('__SourceFile','__SourcePage','User Name','Domain Name','SID','Time Generated','Message','Source')
    for($c=0;$c-lt $headers.Count;$c++){$s.Cells[1,($c+1)].Value=$headers[$c]}
    for($r=0;$r-lt 103;$r++) {
        $time=([datetime]'2026-12-31T23:59:00').AddSeconds(102-$r).ToString('yyyy-MM-dd HH:mm:ss')
        $values=@('fixture.pdf','2','alice','example.local',"S-1-5-21-`n00001",$time,('Message '+($r%2)),'Original Source')
        for($c=0;$c-lt $values.Count;$c++){$s.Cells[($r+2),($c+1)].Value=$values[$c]}
    }
    foreach($row in @(
        @('fixture.pdf','3','Alice','example.local','Another SID','十一月 01,2026 12:00:00 上午','Case-sensitive identity','Original Source'),
        @('fixture.pdf','3','<img src=x onerror="window.injected=1">','example.local','SID','十一月 01,2026 12:00:00 下午','</script><script>window.injected=1</script>','Original Source')
    )) {
        $r=$s.Dimension.End.Row+1
        for($c=0;$c-lt $row.Count;$c++){$s.Cells[$r,($c+1)].Value=$row[$c]}
    }
    for($i=0;$i-lt 11;$i++) {
        $r=$s.Dimension.End.Row+1
        $values=@('fixture.pdf','3',"other-$i",'example.local','SID','2026-01-01 00:00:00','Other message','Original Source')
        for($c=0;$c-lt $values.Count;$c++){$s.Cells[$r,($c+1)].Value=$values[$c]}
    }
    $s=$p.Workbook.Worksheets.Add('Resources')
    $headers=@('__SourceFile','__SourcePage','Reference Resource','Domain Name','Resource   Accessed Time','Message')
    for($c=0;$c-lt $headers.Count;$c++){$s.Cells[1,($c+1)].Value=$headers[$c]}
    $r=2
    foreach($row in @(
        @('resource.pdf','2','server-a','example.local','October 20,2026 01:53:42 PM','Same message'),
        @('resource.pdf','2','server-a','example.local','October 20,2026 01:53:42 PM','Same message'),
        @('resource.pdf','3','server-b','example.local','invalid time','Keep raw text')
    )) {for($c=0;$c-lt $row.Count;$c++){$s.Cells[$r,($c+1)].Value=$row[$c]};$r++}
    $formats=@(
        @('Active Alerts (Trouble)','Source','Time Generated','Alert Message'),
        @('All First Time Resource Access','Reference Resource','Resource Accessed Time','Message'),
        @('All Threshold Breach Activity','Name','Time Generated','Mean Count','Threshold Count','Message'),
        @('All Unusual Time Activity','Name','Time Generated','General Start Time','General End Time','Message'),
        @('First Time Remote Access on Host','Unusual Remote Access From Computer','Resource Accessed Time','Message'),
        @('Unusual Logon Activity Time','User Name','Time Generated','General Start Time','General End Time','Message'),
        @('Unusual Logon Activity Time on Host','Computer Name','Time Generated','General Start Time','General End Time','Message'),
        @('Unusual User Activity Time','User Name','Time Generated','General Start Time','General End Time','Message'),
        @('Unusual Volume of Lockout','User Name','Time Generated','Mean Count','Threshold Count','Message'),
        @('Unusual Volume of Logon Failure','User Name','Time Generated','Mean Count','Threshold Count','Message'),
        @('Unusual Volume of Logon Failure on Host','Computer Name','Time Generated','Mean Count','Threshold Count','Message'),
        @('Unusual Volume of User Management Activity','User Name','Time Generated','Mean Count','Threshold Count','Message')
    )
    for($i=0;$i-lt $formats.Count;$i++) {
        $format=$formats[$i];$sheetName="Format-$i"
        $index.Cells[($i+4),1].Value=$format[0];$index.Cells[($i+4),2].Value=$sheetName
        $s=$p.Workbook.Worksheets.Add($sheetName)
        # Reverse the PDF columns and put an excluded column first to catch positional mappings.
        $headers=@('Domain Name','__SourceFile','__SourcePage')+@($format[($format.Count-1)..1])
        for($c=0;$c-lt $headers.Count;$c++) {
            $name=$headers[$c];$s.Cells[1,($c+1)].Value=$name
            if($name -eq 'Resource Accessed Time'){$s.Cells[1,($c+1)].Value='Resource   Accessed Time'}
            for($r=2;$r-le 3;$r++) {
                $value=switch($name) {
                    $format[1] { 'same-key' }
                    {$_ -in 'Time Generated','Resource Accessed Time'} { '2026-11-01 13:00:00' }
                    default { "shared-$name" }
                }
                $s.Cells[$r,($c+1)].Value=$value
            }
        }
    }
    $p.SaveAs([IO.FileInfo]::new($workbook))
} finally {$p.Dispose()}
& (Join-Path $tool 'New-OfflineReport.ps1') -WorkbookPath $workbook -OutputPath $html -Title 'Fixture <safe>' | Out-Host
& (Join-Path $tool 'New-OfflineReport.ps1') -WorkbookPath $workbook -OutputPath $html -Title 'Replacement fixture' | Out-Host
if([IO.File]::ReadAllText($html) -notmatch '<title>Replacement fixture</title>'){throw 'Existing HTML was not replaced with new report content'}
$savedHash=(Get-FileHash $html).Hash
$outputLock=[IO.File]::Open($html,'Open','Read','None')
$lockRejected=$false
try {
    try { & (Join-Path $tool 'New-OfflineReport.ps1') -WorkbookPath $workbook -OutputPath $html -Title 'Must not overwrite' | Out-Host }
    catch { $lockRejected=$true }
} finally {$outputLock.Dispose()}
if(!$lockRejected -or (Get-FileHash $html).Hash-ne $savedHash){throw 'Failed replacement did not preserve previous HTML'}
if(@(Get-ChildItem $root -Filter 'report-pending-*.html').Count){throw 'Failed replacement left a pending HTML file'}
& (Join-Path $tool 'New-OfflineReport.ps1') -WorkbookPath $workbook -OutputPath $html -Title 'Fixture <safe>' | Out-Host
Write-Output 'PASS: repeated HTML replacement, locked-output preservation and temporary-file cleanup.'
$savedHash=(Get-FileHash $html).Hash
$p=[OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]::new($workbook))
try {
    $s=$p.Workbook.Worksheets['Format-0']
    for($c=1;$c-le $s.Dimension.End.Column;$c++) {
        if($s.Cells[1,$c].Value -eq 'Alert Message'){$s.Cells[1,$c].Value='Missing Message'}
    }
    $p.SaveAs([IO.FileInfo]::new((Join-Path $root 'missing.xlsx')))
} finally {$p.Dispose()}
$missingRejected=$false
try { & (Join-Path $tool 'New-OfflineReport.ps1') -WorkbookPath (Join-Path $root 'missing.xlsx') -OutputPath $html | Out-Host }
catch { $missingRejected=$_.Exception.Message -like '*Active Alerts (Trouble)*Alert Message*' }
if(!$missingRejected -or (Get-FileHash $html).Hash-ne $savedHash){throw 'Missing field did not report context and preserve existing HTML'}
Write-Output 'PASS: missing required field identifies report and column and preserves HTML.'
$source=[IO.File]::ReadAllText($html)
if($source -notmatch '\\u003c/script' -or $source -match '(?i)<(?:script|link)[^>]+(?:src|href)=') {throw 'Unsafe embedded data or external assets'}
if($source -match 'REPORT_DATA|REPORT_CSS|REPORT_JS|REPORT_TITLE'){throw 'Unresolved report template token'}
$browserProcess=$null;$socket=$null
try {
    $profile=Join-Path $root 'browser'
    $browserProcess=Start-Process $Browser -ArgumentList "--headless=new --disable-gpu --no-first-run --no-default-browser-check --remote-debugging-port=0 --remote-allow-origins=* --user-data-dir=`"$profile`" about:blank" -WindowStyle Hidden -RedirectStandardError (Join-Path $root 'browser-error.log') -PassThru
    $portFile=Join-Path $profile 'DevToolsActivePort'
    for($i=0;$i-lt 100 -and !(Test-Path $portFile);$i++){Start-Sleep -Milliseconds 100}
    if(!(Test-Path $portFile)){throw 'Browser did not start'}
    $port=([IO.File]::ReadAllLines($portFile))[0]
    $target=$null
    for($i=0;$i-lt 100;$i++) {
        $targets=Invoke-RestMethod "http://127.0.0.1:$port/json/list"
        $target=$targets | Where-Object type -eq 'page' | Select-Object -First 1
        if($target.webSocketDebuggerUrl){break}
        Start-Sleep -Milliseconds 100
    }
    if(!$target.webSocketDebuggerUrl){throw 'Browser did not expose a page target'}
    $socket=[Net.WebSockets.ClientWebSocket]::new()
    $null=$socket.ConnectAsync([uri]$target.webSocketDebuggerUrl,[Threading.CancellationToken]::None).GetAwaiter().GetResult()
    $script:commandId=0
    function Cdp($Method,$Parameters=@{}) {
        $script:commandId++;$id=$script:commandId
        $bytes=[Text.Encoding]::UTF8.GetBytes((@{id=$id;method=$Method;params=$Parameters}|ConvertTo-Json -Depth 20 -Compress))
        $socket.SendAsync([ArraySegment[byte]]::new($bytes),[Net.WebSockets.WebSocketMessageType]::Text,$true,[Threading.CancellationToken]::None).GetAwaiter().GetResult()
        do {
            $message=[IO.MemoryStream]::new()
            do {
                $buffer=[byte[]]::new(65536);$timeout=[Threading.CancellationTokenSource]::new(30000)
                try{$received=$socket.ReceiveAsync([ArraySegment[byte]]::new($buffer),$timeout.Token).GetAwaiter().GetResult()}finally{$timeout.Dispose()}
                $message.Write($buffer,0,$received.Count)
            }until($received.EndOfMessage)
            $reply=[Text.Encoding]::UTF8.GetString($message.ToArray())|ConvertFrom-Json;$message.Dispose()
            if($reply.method-eq 'Runtime.exceptionThrown'){throw ($reply.params|ConvertTo-Json -Depth 12 -Compress)}
        }until($reply.id-eq $id)
        if($reply.error){throw ($reply.error|ConvertTo-Json -Compress)}
        return $reply.result
    }
    function Evaluate([string]$Code) {
        $answer=Cdp 'Runtime.evaluate' @{expression=$Code;awaitPromise=$true;returnByValue=$true}
        if($answer.exceptionDetails){throw ($answer.exceptionDetails|ConvertTo-Json -Depth 12 -Compress)}
        return $answer.result.value
    }
    $null=Cdp 'Page.enable';$null=Cdp 'Runtime.enable';$null=Cdp 'Network.enable'
    $null=Cdp 'Network.emulateNetworkConditions' @{offline=$true;latency=0;downloadThroughput=0;uploadThroughput=0}
    $null=Cdp 'Emulation.setDeviceMetricsOverride' @{width=1440;height=1050;deviceScaleFactor=1;mobile=$false}
    $null=Cdp 'Page.navigate' @{url=([uri]$html).AbsoluteUri}
    $null=Evaluate '(async()=>{for(let n=0;n<100&&!window.ReportViewer;n++)await new Promise(r=>setTimeout(r,50));if(!window.ReportViewer)throw Error("Report did not initialize");return true})()'
    $checks=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'report-checks.js'))
    Evaluate $checks | ConvertTo-Json -Depth 8 | Out-Host
    if($ActualReport) {
        $null=Cdp 'Page.navigate' @{url=([uri][IO.Path]::GetFullPath($ActualReport)).AbsoluteUri}
        $null=Evaluate '(async()=>{for(let n=0;n<200&&!window.ReportViewer;n++)await new Promise(r=>setTimeout(r,50));if(!window.ReportViewer)throw Error("Actual report did not initialize");return true})()'
        Evaluate '(()=>{const v=ReportViewer;let rows=0,groups=0,invalidTimes=0;for(let i=0;i<v.data.reports.length;i++){v.selectSheet(i);const x=v.getGroups(i);rows+=x.groups.reduce((s,g)=>s+g.rows.length,0);groups+=x.groups.length;invalidTimes+=x.invalidTimes;if(document.querySelectorAll(".group-card").length>12)throw Error("Unbounded group rendering");if(document.querySelectorAll("tbody tr").length>100)throw Error("Unbounded detail rendering");const report=v.data.reports[i];const expected=[report.keyIndex,...report.detailIndices].map(c=>report.columns[c]);const headers=[...document.querySelectorAll("thead th")].map(x=>x.textContent);if(x.groups.length&&JSON.stringify(headers)!==JSON.stringify(expected))throw Error("Incorrect detail columns: "+report.title);for(const g of x.groups){for(let r=1;r<g.rows.length;r++){const a=g.rows[r-1].time,b=g.rows[r].time;if((!a.valid&&b.valid)||(a.valid&&b.valid&&a.key>b.key))throw Error("Incorrect ascending order: "+report.title)}}}v.selectSheet(0);if(invalidTimes)throw Error("Actual report has unconverted times: "+invalidTimes);return {reports:v.data.reports.length,rows,groups,invalidTimes}})()' | ConvertTo-Json | Out-Host
    }
    foreach($width in 1440,900,390) {
        $null=Cdp 'Emulation.setDeviceMetricsOverride' @{width=$width;height=1050;deviceScaleFactor=1;mobile=$false}
        $null=Evaluate 'new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))'
        $layout=Evaluate '({width:innerWidth,scroll:document.documentElement.scrollWidth})'
        if($layout.scroll-gt $layout.width){throw "Horizontal page overflow at $width"}
        $shot=Cdp 'Page.captureScreenshot' @{format='png'}
        [IO.File]::WriteAllBytes((Join-Path $root ("report-$width.png")),[Convert]::FromBase64String($shot.data))
    }
    Write-Output "PASS: offline file, grouping, times, safe data, navigation, pagination and responsive layouts. Screenshots: $root"
} finally {
    if($socket){$socket.Dispose()}
    if($browserProcess -and !$browserProcess.HasExited){Stop-Process -Id $browserProcess.Id -Force}
}

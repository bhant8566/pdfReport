#requires -Version 5.1
[CmdletBinding()]
param([string]$SampleDirectory = "$env:USERPROFILE\Downloads")
$ErrorActionPreference='Stop'
$toolRoot=Split-Path $PSScriptRoot -Parent
$importer=Join-Path $toolRoot 'Import-PdfReports.ps1'
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('pdf-extractor-tests-'+[guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($testRoot) | Out-Null
$checks=[Collections.Generic.List[string]]::new()
function Assert($Value,[string]$Message) { if(!$Value){throw $Message};$checks.Add($Message) }
function Run-Import([string]$InputLocation,[string]$DataLocation,[int]$Page=1,[string]$Script=$importer) {
    & $Script -InputPath $InputLocation -DataRoot $DataLocation -StartPage $Page | Out-Host
    return $LASTEXITCODE
}
Import-Module (Join-Path $toolRoot 'Extractor.Core.psm1') -Force
Initialize-Extractor
function Open-Workbook([string]$Path) { return [OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]::new($Path)) }
function Sheet-Count($Package,[string]$Name) { return $Package.Workbook.Worksheets[$Name].Dimension.End.Row-1 }
function Check-WorkbookOutput($DataLocation) {
    $batch=Get-ChildItem (Join-Path $DataLocation 'Output') -Directory|Sort-Object CreationTime -Descending|Select-Object -First 1
    $batchFiles=@(Get-ChildItem -LiteralPath $batch.FullName -Recurse -File)
    Assert ($batchFiles.Count-eq 1 -and $batchFiles[0].Name-eq 'manifest.json') 'Batch produces only a progress manifest, with no CSV or JSON dataset exports'
    $book=Open-Workbook (Join-Path $DataLocation 'Reports.xlsx')
    try {
        $index=$book.Workbook.Worksheets['_Reports']
        for($r=2;$r-le $index.Dimension.End.Row;$r++) {
            $sheet=$book.Workbook.Worksheets[[string]$index.Cells[$r,2].Value]
            Assert ($sheet -and [int]$index.Cells[$r,3].Value-eq $sheet.Dimension.End.Row-1) 'Saved report row count equals workbook index count'
        }
    } finally { $book.Dispose() }
}
function New-GridPdf([string]$Path,[string]$Title,[object[]]$Pages,[switch]$Cover,[switch]$InvalidLastPage,[switch]$InlineRange,[string]$HeaderRange='From 2026-11-01 to 2026-11-08',[int]$CoverRecordCount=-1) {
    $builder=[UglyToad.PdfPig.Writer.PdfDocumentBuilder]::new()
    $font=$builder.AddStandard14Font([UglyToad.PdfPig.Fonts.Standard14Fonts.Standard14Font]::Helvetica)
    if($Cover) {
        $coverPage=$builder.AddPage([double]600,[double]800)
        $null=$coverPage.AddText('COVER - must not be imported',[double]18,[UglyToad.PdfPig.Core.PdfPoint]::new(40,720),$font)
        if ($CoverRecordCount -ge 0) { $null=$coverPage.AddText(('Number of Records : '+$CoverRecordCount),[double]12,[UglyToad.PdfPig.Core.PdfPoint]::new(40,680),$font) }
    }
    foreach($definition in $Pages) {
        $page=$builder.AddPage([double]600,[double]800)
        if($InlineRange) {
            $null=$page.AddText(($Title+' ('+$HeaderRange+')'),[double]8,[UglyToad.PdfPig.Core.PdfPoint]::new(30,760),$font)
        } else {
            $null=$page.AddText($Title,[double]10,[UglyToad.PdfPig.Core.PdfPoint]::new(30,760),$font)
            $null=$page.AddText($HeaderRange,[double]8,[UglyToad.PdfPig.Core.PdfPoint]::new(390,760),$font)
        }
        $width=540/$definition.Headers.Count
        $allRows=@(@{Values=$definition.Headers})+@($definition.Rows)
        if ($definition.OmitHeader) { $allRows=@($definition.Rows) }
        $top=730
        for($r=0;$r-lt $allRows.Count;$r++) {
            $height=if($r-eq 0 -and !$definition.OmitHeader){35}else{95}
            for($c=0;$c-lt $definition.Headers.Count;$c++) {
                $x=30+$c*$width
                $null=$page.DrawRectangle([UglyToad.PdfPig.Core.PdfPoint]::new($x,$top-$height),[double]$width,[double]$height,[double]0.5)
                $lineIndex=0
                foreach($line in ([string]$allRows[$r].Values[$c] -split "`n")) {
                    $null=$page.AddText($line,[double]9,[UglyToad.PdfPig.Core.PdfPoint]::new($x+4,$top-14-$lineIndex*12),$font)
                    $lineIndex++
                }
            }
            $top-=$height
        }
        $null=$page.AddText('Supplier footer - exclude',[double]9,[UglyToad.PdfPig.Core.PdfPoint]::new(30,20),$font)
    }
    if($InvalidLastPage){$p=$builder.AddPage([double]600,[double]800);$null=$p.AddText('Readable page without a table',[double]12,[UglyToad.PdfPig.Core.PdfPoint]::new(30,700),$font)}
    [IO.File]::WriteAllBytes($Path,$builder.Build())
    $builder.Dispose()
}

$inputFolder=Join-Path $testRoot 'samples'
[IO.Directory]::CreateDirectory($inputFolder)|Out-Null
$names=@('Unusual Volume of Logon Failure','Unusual Logon Activity Time','All Unusual Time Activity')
$sourceHashes=@{}
foreach($name in $names) {
    $path=Join-Path $SampleDirectory ($name+'.pdf')
    if(!(Test-Path -LiteralPath $path)){throw "Missing provided sample: $path. Supply -SampleDirectory."}
    $sourceHashes[$path]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    Copy-Item -LiteralPath $path -Destination $inputFolder
}
$data=Join-Path $testRoot 'history'
$oldOutput=Join-Path $data 'Output/previous-exports'
[IO.Directory]::CreateDirectory($oldOutput)|Out-Null
$oldCsv=Join-Path $oldOutput 'existing.csv'; $oldJson=Join-Path $oldOutput 'existing.json'
[IO.File]::WriteAllText($oldCsv,'previous CSV data')
[IO.File]::WriteAllText($oldJson,'{"previous":"JSON data"}')
$oldCsvHash=(Get-FileHash $oldCsv).Hash; $oldJsonHash=(Get-FileHash $oldJson).Hash
Assert ((Run-Import $inputFolder $data)-eq 0) 'Three provided samples import successfully'
Assert ((Get-FileHash $oldCsv).Hash-eq $oldCsvHash -and (Get-FileHash $oldJson).Hash-eq $oldJsonHash) 'Existing CSV and JSON exports remain unchanged'
$lastBatch=Get-ChildItem (Join-Path $data 'Output') -Directory|Sort-Object CreationTime -Descending|Select-Object -First 1
$progressManifest=Get-Content (Join-Path $lastBatch.FullName 'manifest.json') -Raw -Encoding UTF8|ConvertFrom-Json
Assert ($progressManifest.Current.Page-eq 2 -and $progressManifest.Current.TotalPages-eq 2 -and $progressManifest.Current.Phase-eq 'Finished') 'Manifest records page totals, completion phase and current file'
Assert ($progressManifest.Current.ElapsedSeconds-ge 0 -and $progressManifest.Current.UpdatedAt) 'Manifest includes elapsed time and latest heartbeat'
$events=[Collections.Generic.List[object]]::new()
$null=Read-ReportFile (Join-Path $inputFolder ($names[0]+'.pdf')) 1 'progress-test' 'now' {param($page,$total,$rows,$phase) $events.Add([pscustomobject]@{Page=$page;Total=$total;Rows=$rows;Phase=$phase})}
Assert ($events.Count-eq 4 -and $events[0].Page-eq 1 -and $events[0].Phase-eq 'ParsingPage' -and $events[0].Rows-eq 0) 'Progress callback runs before parsing first page'
Assert ($events[3].Page-eq 2 -and $events[3].Rows-eq 32 -and $events[3].Phase-eq 'PageComplete') 'Progress callback reports every completed page and cumulative row count'
$workbook=Join-Path $data 'Reports.xlsx'
$book=Open-Workbook $workbook
try {
    Assert ($book.Workbook.Worksheets.Count-eq 5) 'Three report sheets and two index sheets are created'
    Assert ((Sheet-Count $book $names[0])-eq 32) 'Failure report has 32 complete records'
    Assert ((Sheet-Count $book $names[1])-eq 32) 'Logon-time report has 32 complete records'
    Assert ((Sheet-Count $book $names[2])-eq 18) 'All-time report has 18 complete records'
    foreach($name in $names) {
        $sheet=$book.Workbook.Worksheets[$name]
        $headers=@((1..$sheet.Dimension.End.Column)|ForEach-Object{[string]$sheet.Cells[1,$_].Value})
        Assert ($headers -ccontains 'Message') 'Repeated letters in Message header remain intact'
        Assert ($headers.Count-eq $(if($name-eq $names[2]){10}else{12})) 'Original 8/10 columns plus only 2 metadata columns are preserved'
        Assert ($sheet.Cells[2,2].Value-eq '1') 'Source physical page is preserved'
        Assert ((@($headers | Where-Object {$_ -like '__*'}) -join ',')-ceq '__SourceFile,__SourcePage') 'Only requested metadata is written to report details'
        Assert ($sheet.Cells[2,1].Value-ceq ($name+'.pdf')) 'Source filename is retained in report details'
    }
    $failure=$book.Workbook.Worksheets[$names[0]]
    Assert ($failure.Cells[3,4].Value -ceq "S-1-5-21-81404079-`n1603697135-622671684-`n65775") 'Long wrapped SID stays complete in one cell'
    Assert (($failure.Cells[2,7].Value -replace '\s','') -ceq '十月31,202611:59:46下午') 'Wrapped Chinese date retains all characters'
    $message="10+ number of Logon`nFailure Activity occured`nfor Bad UserNames within`n11PM - 12AM. Usual`naverage is 0, Threshold`ncalculated is 10. Anomaly`ncategory:Unusual Activity`n-Logon Failure Count`n(Based on User)"
    Assert ($failure.Cells[2,11].Value-ceq $message) 'Multiline message and spaces match the original PDF'
    Assert ($failure.Cells[2,3].Value-ceq 'Bad UserNames') 'User names keep original spaces'
} finally {$book.Dispose()}
Check-WorkbookOutput $data
Assert ((Run-Import $inputFolder $data)-eq 0) 'Repeated import succeeds without deduplication'
$book=Open-Workbook $workbook
try {
    Assert ((Sheet-Count $book $names[0])-eq 64 -and (Sheet-Count $book $names[2])-eq 36) 'History appends all repeated records'
    Assert ((Sheet-Count $book '_Imports')-eq 6) 'Import log records both batches'
} finally {$book.Dispose()}
Assert (Test-Path (Join-Path $data 'Reports.previous.xlsx')) 'Previous workbook backup exists'
Check-WorkbookOutput $data
$defaultData=Join-Path $testRoot 'default-page'
Assert ((Run-Import (Join-Path $inputFolder ($names[0]+'.pdf')) $defaultData 2)-eq 0) 'Default start-page behavior works on sample subset'
$book=Open-Workbook (Join-Path $defaultData 'Reports.xlsx')
try {Assert ((Sheet-Count $book $names[0])-eq 16 -and $book.Workbook.Worksheets[$names[0]].Cells[2,2].Value-eq '2') 'StartPage 2 skips physical page 1'} finally {$book.Dispose()}

$synthetic=Join-Path $testRoot 'evolution.pdf'
$first=@{Headers=@('Identifier','Value','Message');Rows=@(@{Values=@('001','0',"Hello, `"world`"`nMore text")})}
$second=@{Headers=@('Message','Identifier','Value','Extra');Rows=@(@{Values=@('Second message','002','12.50','New column')})}
$longTitle='This report title exceeds thirty one characters and must use a stable mapping'
New-GridPdf $synthetic $longTitle @($first,$second) -Cover
$evolution=Join-Path $testRoot 'evolution'
Assert ((Run-Import $synthetic $evolution 2)-eq 0) 'Real cover and changing per-page headers import successfully'
$book=Open-Workbook (Join-Path $evolution 'Reports.xlsx')
try {
    $sheetName=[string]$book.Workbook.Worksheets['_Reports'].Cells[2,2].Value
    $sheet=$book.Workbook.Worksheets[$sheetName]
    Assert ($sheetName.Length-le 31 -and (Sheet-Count $book $sheetName)-eq 2) 'Long report title maps to a short worksheet name'
    Assert ($sheet.Cells[2,3].Value-ceq '001' -and $sheet.Cells[3,3].Value-ceq '002') 'Reordered headers map by name and leading zero survives'
    Assert (!$sheet.Cells[2,6].Value -and $sheet.Cells[3,6].Value-ceq 'New column') 'New column leaves old rows empty'
    Assert ($sheet.Cells[2,5].Value-ceq "Hello, `"world`"`nMore text") 'Comma, quotation marks and newline remain literal cell data'
} finally {$book.Dispose()}
Check-WorkbookOutput $evolution
Assert ((Run-Import $synthetic $evolution 2)-eq 0) 'Long-name worksheet mapping remains stable on later imports'
$book=Open-Workbook (Join-Path $evolution 'Reports.xlsx')
try {Assert ($book.Workbook.Worksheets.Count-eq 3 -and (Sheet-Count $book $sheetName)-eq 4) 'Second long-title import appends to existing worksheet'} finally {$book.Dispose()}

$legacyData=Join-Path $testRoot 'legacy-history'
[IO.Directory]::CreateDirectory($legacyData)|Out-Null
$legacyPath=Join-Path $legacyData 'Reports.xlsx'
Copy-Item -LiteralPath (Join-Path $evolution 'Reports.xlsx') -Destination $legacyPath
$legacyBook=Open-Workbook $legacyPath
try {
    $legacySheet=$legacyBook.Workbook.Worksheets[$sheetName]
    $legacyFields=@('__ReportType','__SourcePath','__SourceRow','__ReportRange','__ImportedAt','__BatchId','__FileHash')
    for($c=$legacyFields.Count-1;$c-ge 0;$c--) {
        $legacySheet.InsertColumn(1,1)
        $legacySheet.Cells[1,1].Value=$legacyFields[$c]
        for($r=2;$r-le $legacySheet.Dimension.End.Row;$r++) { $legacySheet.Cells[$r,1].Value='Legacy metadata' }
    }
    $sourceColumn=$legacySheet.Dimension.End.Column+1
    $legacySheet.Cells[1,$sourceColumn].Value='Source'
    for($r=2;$r-le $legacySheet.Dimension.End.Row;$r++) { $legacySheet.Cells[$r,$sourceColumn].Value='Original PDF Source' }
    $legacyBook.Save()
} finally {$legacyBook.Dispose()}
$legacyHash=(Get-FileHash $legacyPath).Hash
Assert ((Run-Import $synthetic $legacyData 2)-eq 0) 'Existing workbook with legacy metadata migrates during atomic per-file save'
$legacyBook=Open-Workbook $legacyPath
try {
    $legacySheet=$legacyBook.Workbook.Worksheets[$sheetName]
    $legacyHeaders=@((1..$legacySheet.Dimension.End.Column)|ForEach-Object {[string]$legacySheet.Cells[1,$_].Value})
    Assert ((@($legacyHeaders|Where-Object {$_ -like '__*'})-join ',')-ceq '__SourceFile,__SourcePage') 'Migration removes all seven old metadata fields'
    Assert ((Sheet-Count $legacyBook $sheetName)-eq 6 -and $legacySheet.Cells[2,3].Value-ceq '001' -and $legacySheet.Cells[2,5].Value-ceq "Hello, `"world`"`nMore text") 'Migration preserves historical rows, leading zeros and multiline text while appending new rows'
    $sourceColumn=[array]::IndexOf($legacyHeaders,'Source')+1
    Assert ($sourceColumn-gt 0 -and $legacySheet.Cells[2,$sourceColumn].Value-ceq 'Original PDF Source') 'PDF Source field survives removal of tool metadata'
} finally {$legacyBook.Dispose()}
Assert ((Get-FileHash (Join-Path $legacyData 'Reports.previous.xlsx')).Hash-eq $legacyHash) 'Migration backup preserves exact pre-migration workbook'
Check-WorkbookOutput $legacyData

$mixed=Join-Path $testRoot 'mixed'
[IO.Directory]::CreateDirectory($mixed)|Out-Null
Copy-Item -LiteralPath $synthetic -Destination $mixed
[IO.File]::WriteAllText((Join-Path $mixed 'broken.pdf'),'not a PDF')
New-GridPdf (Join-Path $mixed 'partial.pdf') 'Partial report must not be committed' @($first) -Cover -InvalidLastPage
$mixedData=Join-Path $testRoot 'mixed-data'
Assert ((Run-Import $mixed $mixedData 2)-eq 1) 'One valid and two invalid files return partial-success exit code'
$book=Open-Workbook (Join-Path $mixedData 'Reports.xlsx')
try {
    Assert ($book.Workbook.Worksheets.Count-eq 3) 'Failed PDF with a valid first table writes no partial report sheet'
    Assert ((Sheet-Count $book '_Imports')-eq 3) 'Import log includes per-file failures'
} finally {$book.Dispose()}
$before=(Get-FileHash $workbook -Algorithm SHA256).Hash
$handle=[IO.File]::Open($workbook,'Open','ReadWrite','None')
try {Assert ((Run-Import $inputFolder $data)-eq 2) 'Locked workbook returns storage failure'} finally {$handle.Dispose()}
Assert ((Get-FileHash $workbook -Algorithm SHA256).Hash-eq $before) 'Locked workbook remains byte-for-byte unchanged'
$allBad=Join-Path $testRoot 'all-bad'
Assert ((Run-Import (Join-Path $mixed 'partial.pdf') $allBad 2)-eq 1) 'All-invalid batch fails'
Assert (!(Test-Path (Join-Path $allBad 'Reports.xlsx'))) 'All-invalid batch does not create history workbook'
$duplicateHeader=Join-Path $testRoot 'duplicate-header.pdf'
New-GridPdf $duplicateHeader 'Ambiguous headers' @(@{Headers=@('Name','name');Rows=@(@{Values=@('Keep A','Keep B')})}) -Cover
$duplicateData=Join-Path $testRoot 'duplicate-data'
Assert ((Run-Import $duplicateHeader $duplicateData 2)-eq 1) 'Case-insensitive duplicate headers fail rather than overwrite fields'
Assert (!(Test-Path (Join-Path $duplicateData 'Reports.xlsx'))) 'Ambiguous header produces no partial workbook'

$collisions=Join-Path $testRoot 'collisions'
[IO.Directory]::CreateDirectory($collisions)|Out-Null
New-GridPdf (Join-Path $collisions 'a.pdf') 'A<B' @($first) -Cover
New-GridPdf (Join-Path $collisions 'b.pdf') 'A_B' @($first) -Cover
$collisionData=Join-Path $testRoot 'collision-data'
Assert ((Run-Import $collisions $collisionData 2)-eq 0) 'Reports with colliding sanitized names import successfully'
$book=Open-Workbook (Join-Path $collisionData 'Reports.xlsx')
try {Assert ($book.Workbook.Worksheets.Count-eq 4) 'Distinct report titles retain separate worksheets after sanitization'} finally {$book.Dispose()}
Check-WorkbookOutput $collisionData

$inlineInput=Join-Path $testRoot 'inline-range'
[IO.Directory]::CreateDirectory($inlineInput)|Out-Null
New-GridPdf (Join-Path $inlineInput 'october.pdf') 'Active Alerts (Trouble)' @($first) -Cover -InlineRange -HeaderRange 'From 2026-10-01 To 2026-11-01'
New-GridPdf (Join-Path $inlineInput 'november.pdf') 'Active Alerts (Trouble)' @($first) -Cover -InlineRange -HeaderRange 'From 2026-11-01 To 2026-12-01'
$inlineData=Join-Path $testRoot 'inline-data'
Assert ((Run-Import $inlineInput $inlineData 2)-eq 0) 'Reports with parenthesized inline date ranges import successfully'
$book=Open-Workbook (Join-Path $inlineData 'Reports.xlsx')
try {
    Assert ($book.Workbook.Worksheets.Count-eq 3 -and (Sheet-Count $book 'Active Alerts (Trouble)')-eq 2) 'Changing inline report period appends to same report-type worksheet'
    $sheet=$book.Workbook.Worksheets['Active Alerts (Trouble)']
    Assert ($sheet.Cells[2,1].Value -ne $sheet.Cells[3,1].Value -and $sheet.Name -eq 'Active Alerts (Trouble)') 'Changing date ranges do not become extra detail columns or part of report title'
} finally {$book.Dispose()}
Check-WorkbookOutput $inlineData

$continuationPdf=Join-Path $testRoot 'omitted-headers.pdf'
$continued=@{Headers=$first.Headers;OmitHeader=$true;Rows=@(@{Values=@('002','Repeated','Data with no repeated header')},@{Values=@('003','Repeated','Another record')})}
New-GridPdf $continuationPdf 'Continuation report' @($first,$continued,$first) -Cover -CoverRecordCount 4
$continuationData=Join-Path $testRoot 'continuation-data'
Assert ((Run-Import $continuationPdf $continuationData 2)-eq 0) 'Headerless continuation page imports using verified prior geometry'
$book=Open-Workbook (Join-Path $continuationData 'Reports.xlsx')
try {
    $sheet=$book.Workbook.Worksheets['Continuation report']
    Assert (($sheet.Dimension.End.Row-1)-eq 4 -and $sheet.Cells[3,3].Value-ceq '002' -and $sheet.Cells[4,3].Value-ceq '003') 'Headerless page preserves its first record and repeated values'
    Assert ($sheet.Cells[3,2].Value-eq '3') 'Continuation rows retain physical source page'
} finally {$book.Dispose()}
Check-WorkbookOutput $continuationData
$mismatchPdf=Join-Path $testRoot 'count-mismatch.pdf'
New-GridPdf $mismatchPdf 'Count mismatch' @($first) -Cover -CoverRecordCount 2
$mismatchData=Join-Path $testRoot 'count-mismatch-data'
Assert ((Run-Import $mismatchPdf $mismatchData 2)-eq 1) 'Cover count mismatch rejects incomplete extraction'
Assert (!(Test-Path (Join-Path $mismatchData 'Reports.xlsx'))) 'Count mismatch never writes a partial workbook'

$checkpointTool=Join-Path $testRoot 'checkpoint-tool'
[IO.Directory]::CreateDirectory($checkpointTool)|Out-Null
foreach($component in 'Import-PdfReports.ps1','Extractor.Core.psm1','lib') { Copy-Item -LiteralPath (Join-Path $toolRoot $component) -Destination $checkpointTool -Recurse }
$probe=@'
    if ([IO.Path]::GetFileName($Path) -like 'b-*.pdf') {
        $probeWorkbook=Join-Path $env:EXTRACTOR_CHECKPOINT_TEST_DATA 'Reports.xlsx'
        $probeBook=[OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]::new($probeWorkbook))
        try {
            $probeName=[string]$probeBook.Workbook.Worksheets['_Reports'].Cells[2,2].Value
            if ($probeBook.Workbook.Worksheets[$probeName].Dimension.End.Row-ne 3 -or $probeBook.Workbook.Worksheets['_Imports'].Dimension.End.Row-ne 2) { throw 'First PDF was not independently committed before second PDF opened.' }
        } finally { $probeBook.Dispose() }
        [IO.File]::WriteAllText((Join-Path $env:EXTRACTOR_CHECKPOINT_TEST_DATA 'checkpoint-observed.txt'),(Get-FileHash $probeWorkbook -Algorithm SHA256).Hash)
        if ([IO.Path]::GetFileName($Path)-eq 'b-locked.pdf') { $script:probeLock=[IO.File]::Open($probeWorkbook,'Open','ReadWrite','None') }
    }
'@
$modulePath=Join-Path $checkpointTool 'Extractor.Core.psm1'
$moduleSource=[IO.File]::ReadAllText($modulePath)
$functionStart='function Read-ReportFile([string]$Path, [int]$StartPage, [string]$BatchId, [string]$ImportedAt, [scriptblock]$ProgressCallback) {'
Assert ($moduleSource.Contains($functionStart)) 'Checkpoint probe targets the PDF reader in an isolated tool copy'
[IO.File]::WriteAllText($modulePath,$moduleSource.Replace($functionStart,($functionStart+"`n"+$probe+"`n")),[Text.UTF8Encoding]::new($true))
$previousProbeData=$env:EXTRACTOR_CHECKPOINT_TEST_DATA
try {
    foreach($scenario in 'verify','locked') {
        $probeInput=Join-Path $testRoot ('checkpoint-input-'+$scenario)
        $probeData=Join-Path $testRoot ('checkpoint-data-'+$scenario)
        [IO.Directory]::CreateDirectory($probeInput)|Out-Null
        Copy-Item -LiteralPath $synthetic -Destination (Join-Path $probeInput 'a-first.pdf')
        Copy-Item -LiteralPath $synthetic -Destination (Join-Path $probeInput ('b-'+$scenario+'.pdf'))
        $env:EXTRACTOR_CHECKPOINT_TEST_DATA=$probeData
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $checkpointTool 'Import-PdfReports.ps1') -InputPath $probeInput -DataRoot $probeData | Out-Host
        $probeExit=$LASTEXITCODE
        Assert (Test-Path (Join-Path $probeData 'checkpoint-observed.txt')) 'First PDF is saved and readable before parsing second PDF'
        $probeBatch=Get-ChildItem (Join-Path $probeData 'Output') -Directory | Select-Object -First 1
        $probeManifest=Get-Content (Join-Path $probeBatch.FullName 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $probeBook=Open-Workbook (Join-Path $probeData 'Reports.xlsx')
        try {
            $probeName=[string]$probeBook.Workbook.Worksheets['_Reports'].Cells[2,2].Value
            if($scenario-eq 'verify') {
                Assert ($probeExit-eq 0 -and (Sheet-Count $probeBook $probeName)-eq 4 -and (Sheet-Count $probeBook '_Imports')-eq 2) 'Per-file saves append each PDF and import log exactly once'
                $backup=Open-Workbook (Join-Path $probeData 'Reports.previous.xlsx')
                try { Assert ((Sheet-Count $backup $probeName)-eq 2) 'Previous backup contains state before last PDF checkpoint' } finally {$backup.Dispose()}
            } else {
                Assert ($probeExit-eq 2 -and (Sheet-Count $probeBook $probeName)-eq 2 -and (Sheet-Count $probeBook '_Imports')-eq 1) 'Later storage failure preserves earlier PDF records and import log'
                Assert ($probeManifest.Status-eq 'Failed' -and $probeManifest.WorkbookCommitted -and $probeManifest.Files[0].Status-eq 'Success' -and $probeManifest.Files[1].Status-eq 'Failed' -and $probeManifest.Files[1].Rows-eq 0) 'Storage failure reports saved and unsaved files accurately'
                Assert ((Get-FileHash (Join-Path $probeData 'Reports.xlsx')).Hash-eq [IO.File]::ReadAllText((Join-Path $probeData 'checkpoint-observed.txt'))) 'Failed later save leaves prior checkpoint byte-for-byte unchanged'
                Assert (@(Get-ChildItem $probeData -Filter 'pending-*.xlsx').Count-eq 0) 'Failed checkpoint cleans up its temporary workbook'
            }
        } finally {$probeBook.Dispose()}
    }
} finally { $env:EXTRACTOR_CHECKPOINT_TEST_DATA=$previousProbeData }

$portable=Join-Path $testRoot 'portable/PdfReportExtractor'
[IO.Directory]::CreateDirectory((Split-Path $portable -Parent))|Out-Null
Copy-Item -LiteralPath $toolRoot -Destination $portable -Recurse
$offlineData=Join-Path $testRoot 'offline-data'
$offlineCheck=Join-Path $testRoot 'Run-Offline.ps1'
$offlineSource=@'
param($Tool,$InputPdf,$DataDirectory)
$ErrorActionPreference='Stop'
$env:PSModulePath=''
function global:Invoke-WebRequest { throw 'Network must not be used by the importer' }
function global:Invoke-RestMethod { throw 'Network must not be used by the importer' }
& $Tool -InputPath $InputPdf -DataRoot $DataDirectory -StartPage 2
exit $LASTEXITCODE
'@
[IO.File]::WriteAllText($offlineCheck,$offlineSource,[Text.UTF8Encoding]::new($true))
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $offlineCheck -Tool (Join-Path $portable 'Import-PdfReports.ps1') -InputPdf $synthetic -DataDirectory $offlineData | Out-Host
Assert ($LASTEXITCODE-eq 0 -and (Test-Path (Join-Path $offlineData 'Reports.xlsx'))) 'Moved tool works in fresh PowerShell with empty module search path and blocked web cmdlets'
$defaultTool=Join-Path $testRoot 'default-tool'
[IO.Directory]::CreateDirectory((Join-Path $defaultTool 'Input'))|Out-Null
foreach($component in 'Import-PdfReports.ps1','Extractor.Core.psm1','lib') { Copy-Item -LiteralPath (Join-Path $toolRoot $component) -Destination $defaultTool -Recurse }
Copy-Item -LiteralPath $synthetic -Destination (Join-Path $defaultTool 'Input')
Push-Location $testRoot
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $defaultTool 'Import-PdfReports.ps1') | Out-Host
    Assert ($LASTEXITCODE-eq 0 -and (Test-Path (Join-Path $defaultTool 'Data/Reports.xlsx'))) 'No-argument launch uses Input and Data beside script even from another working directory'
} finally { Pop-Location }
foreach($path in $sourceHashes.Keys) {Assert ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash-eq $sourceHashes[$path]) 'Original sample PDF remains unchanged'}
Write-Output ("PASS: {0} checks. Temporary artifacts: {1}" -f $checks.Count,$testRoot)
exit 0

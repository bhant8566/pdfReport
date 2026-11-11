#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InputDirectory = '',
    [string]$ExistingDataRoot = ''
)
$ErrorActionPreference='Stop'
$toolRoot=Split-Path $PSScriptRoot -Parent
if (!$InputDirectory) { $InputDirectory=Join-Path $toolRoot 'Input' }
Import-Module (Join-Path $toolRoot 'Extractor.Core.psm1') -Force
Initialize-Extractor
$files=@(Get-ChildItem -LiteralPath $InputDirectory -File -Filter '*.pdf' | Sort-Object Name)
if(!$files.Count){throw 'No PDF files to validate.'}
if(!$ExistingDataRoot) {
    $validationRoot=Join-Path ([IO.Path]::GetTempPath()) ('pdf-extractor-input-check-'+[guid]::NewGuid().ToString('N'))
    $ExistingDataRoot=Join-Path $validationRoot 'Data'
    & (Join-Path $toolRoot 'Import-PdfReports.ps1') -InputPath $InputDirectory -DataRoot $ExistingDataRoot
    if($LASTEXITCODE-ne 0){throw "Batch import failed. Inspect $ExistingDataRoot"}
}
$batch=Get-ChildItem (Join-Path $ExistingDataRoot 'Output') -Directory | Sort-Object CreationTime -Descending | Select-Object -First 1
$manifest=Get-Content (Join-Path $batch.FullName 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if($manifest.Status-ne 'Success' -or !$manifest.WorkbookCommitted){throw 'Latest batch did not fully commit successfully.'}
if($manifest.Files.Count-ne $files.Count){throw 'Input count does not match latest batch.'}
$batchFiles=@(Get-ChildItem -LiteralPath $batch.FullName -Recurse -File)
if($batchFiles.Count-ne 1 -or $batchFiles[0].Name-ne 'manifest.json'){throw 'Batch should only contain its progress manifest, without CSV or JSON dataset exports.'}
$book=[OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]::new((Join-Path $ExistingDataRoot 'Reports.xlsx')))
$summary=[Collections.Generic.List[object]]::new()
try {
    foreach($file in $files) {
        $document=[UglyToad.PdfPig.PdfDocument]::Open($file.FullName)
        try {
            $cover=@($document.GetPage(1).GetWords() | ForEach-Object Text) -join ' '
            $count=[regex]::Match($cover,'(?i)Number\s+of\s+Records\s*:\s*(?<count>[\d,]+)')
            if(!$count.Success){throw "No reference record count on cover: $($file.Name)"}
            $expected=[int]$count.Groups['count'].Value.Replace(',','')
            $log=@($manifest.Files | Where-Object SourceFile -eq $file.Name)
            if($log.Count-ne 1 -or $log[0].Status-ne 'Success' -or $log[0].Rows-ne $expected){throw "Cover count mismatch: $($file.Name)"}
            $reportIndex=$book.Workbook.Worksheets['_Reports']
            $report=$null
            for($ir=2;$ir-le $reportIndex.Dimension.End.Row;$ir++) {
                $sheet=$book.Workbook.Worksheets[[string]$reportIndex.Cells[$ir,2].Value]
                if($sheet.Dimension.End.Row-ge 2 -and [string]$sheet.Cells[2,1].Value-ceq $file.Name){
                    $report=[pscustomobject]@{Title=[string]$reportIndex.Cells[$ir,1].Value;Name=$sheet.Name;Rows=[int]$reportIndex.Cells[$ir,3].Value};break
                }
            }
            if(!$report){throw "Missing report worksheet: $($file.Name)"}
            $sheet=$book.Workbook.Worksheets[$report.Name]
            $headers=@(for($c=1;$c-le $sheet.Dimension.End.Column;$c++){[string]$sheet.Cells[1,$c].Value})
            if ((@($headers|Where-Object {$_ -like '__*'})-join ',')-cne '__SourceFile,__SourcePage') {throw "Unexpected metadata columns: $($file.Name)"}
            $index=@{};for($c=0;$c-lt $headers.Count;$c++){$index[$headers[$c]]=$c+1}
            $actual=$sheet.Dimension.End.Row-1
            if($actual-ne $expected -or $report.Rows-ne $expected){throw "Worksheet count mismatch: $($file.Name)"}
            Write-Host ("Checking {0}: {1} rows, {2} source columns" -f $file.Name,$expected,($headers.Count-2))
            for($r=2;$r-le $sheet.Dimension.End.Row;$r++) {
                $page=[int]$sheet.Cells[$r,$index['__SourcePage']].Value
                if($page-lt 2 -or $page-gt $document.NumberOfPages){throw "Invalid source page: $($file.Name), row $r"}
                if([string]$sheet.Cells[$r,$index['__SourceFile']].Value-cne $file.Name){throw 'Source filename mismatch.'}
            }
            $summary.Add([pscustomobject]@{File=$file.Name;ReportTitle=$report.Title;Pages=$document.NumberOfPages;Columns=$headers.Count-2;Rows=$actual;CoverRows=$expected;Sheet=$report.Name})
        } finally {$document.Dispose()}
    }
    if($book.Workbook.Worksheets.Count-ne $files.Count+2){throw 'Unexpected sheet count; expected one sheet per input plus two indexes.'}
} finally {$book.Dispose()}
$summary | Format-Table File,Pages,Columns,Rows,CoverRows -AutoSize
Write-Output ("PASS: {0} reports, {1} PDF pages, {2} rows. Cover counts, workbook indexes and source pages match. No dataset exports. Data: {3}" -f $summary.Count,($summary|Measure-Object Pages -Sum).Sum,($summary|Measure-Object Rows -Sum).Sum,$ExistingDataRoot)

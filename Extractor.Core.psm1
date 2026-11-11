$ErrorActionPreference = 'Stop'
$script:MetaColumns = @('__SourceFile','__SourcePage')
$script:LegacyMetaColumns = @('__ReportType','__SourcePath','__SourceRow','__ReportRange','__ImportedAt','__BatchId','__FileHash')
$script:IndexColumns = @('ReportTitle','SheetName','RowCount','UpdatedAt')
$script:ImportColumns = @('BatchId','ImportedAt','SourceFile','SourcePath','SHA256','StartPage','Status','Rows','Error')

function Initialize-Extractor {
    $lib = Join-Path $PSScriptRoot 'lib'
    foreach ($name in @('System.Runtime.CompilerServices.Unsafe','System.Numerics.Vectors','System.Buffers','System.Memory','Microsoft.Bcl.HashCode','System.ValueTuple')) {
        $file = Join-Path $lib ($name + '.dll')
        if (!(Test-Path -LiteralPath $file)) { throw "缺少離線函式庫：$file。請重新取得完整工具資料夾。" }
        [void][Reflection.Assembly]::LoadFrom($file)
    }
    foreach ($file in Get-ChildItem -LiteralPath $lib -Filter 'UglyToad*.dll') { [void][Reflection.Assembly]::LoadFrom($file.FullName) }
    [void][Reflection.Assembly]::LoadFrom((Join-Path $lib 'Tabula.dll'))
    [void][Reflection.Assembly]::LoadFrom((Join-Path $lib 'PdfReportExtractor.TableText.dll'))
    Import-Module -Name (Join-Path $lib 'ImportExcel/ImportExcel.psd1') -ErrorAction Stop
}

function Get-TextHash([string]$Text) {
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant() }
    finally { $hasher.Dispose() }
}

function Get-CellText($Cell, [object[]]$Letters) {
    # Tabula supplies cell geometry only. Reconstruct original PdfPig glyphs:
    # Tabula's text merging can discard repeated letters and move spaces.
    $inside = @($Letters.Where({ $_.StartBaseLine.X -ge $Cell.Left -and $_.StartBaseLine.X -lt $Cell.Right }))
    $ordered = @($inside | Sort-Object @{Expression={$_.StartBaseLine.Y};Descending=$true})
    $lines = [Collections.Generic.List[string]]::new()
    $lineLetters = [Collections.Generic.List[object]]::new()
    $baseline = [double]::NaN
    foreach ($letter in $ordered) {
        if (![double]::IsNaN($baseline) -and [Math]::Abs($letter.StartBaseLine.Y - $baseline) -gt 0.8) {
            $lines.Add(((@($lineLetters | Sort-Object { $_.StartBaseLine.X } | ForEach-Object Value)) -join '').TrimEnd())
            $lineLetters.Clear()
        }
        if (!$lineLetters.Count) { $baseline = [double]$letter.StartBaseLine.Y }
        $lineLetters.Add($letter)
    }
    if ($lineLetters.Count) { $lines.Add(((@($lineLetters | Sort-Object { $_.StartBaseLine.X } | ForEach-Object Value)) -join '').TrimEnd()) }
    return ($lines -join "`n").Trim()
}

function Get-ReportHeader($Page, [double]$TableTop) {
    $words = @($Page.GetWords() | Where-Object { $_.BoundingBox.Bottom -ge $TableTop })
    if (!$words.Count) { throw '找不到表格上方的報告小標題。' }
    $anchor = $words | Sort-Object @{Expression={$_.BoundingBox.Left}},@{Expression={$_.BoundingBox.Bottom};Descending=$true} | Select-Object -First 1
    $line = @($words | Where-Object { [Math]::Abs($_.BoundingBox.Bottom - $anchor.BoundingBox.Bottom) -lt 2 } | Sort-Object { $_.BoundingBox.Left })
    $title = (@($line | ForEach-Object Text) -join ' ').Trim()
    $range = ''
    $separators = [regex]::Matches($title,'(?i)\s+(?=\(?From\b)')
    if ($separators.Count) {
        $separator=$separators[$separators.Count-1]
        $candidate=$title.Substring($separator.Index).Trim()
        if ($candidate -match '(?i)\bTo\b' -and $candidate -match '\d') {
            $range=$candidate
            $title=$title.Substring(0,$separator.Index).Trim()
        }
    }
    if (!$title) { throw '報告小標題為空，無法決定 Excel 分頁。' }
    return [pscustomobject]@{Title=$title;Range=$range}
}

function Read-ReportFile([string]$Path, [int]$StartPage, [string]$BatchId, [string]$ImportedAt, [scriptblock]$ProgressCallback) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { $hash = ([BitConverter]::ToString($hasher.ComputeHash($bytes))).Replace('-','').ToLowerInvariant() }
    finally { $hasher.Dispose() }
    $options = [UglyToad.PdfPig.ParsingOptions]::new()
    $options.ClipPaths = $true
    $document = [UglyToad.PdfPig.PdfDocument]::Open($bytes,$options)
    $records = [Collections.Generic.List[object]]::new()
    $schemas = [Collections.Generic.List[object]]::new()
    $knownHeaders = @{}
    try {
        if ($StartPage -gt $document.NumberOfPages) { throw "只有 $($document.NumberOfPages) 頁，起始頁 $StartPage 之後沒有資料。" }
        $expectedCount = $null
        if ($StartPage -eq 2) {
            $coverText = @($document.GetPage(1).GetWords() | ForEach-Object Text) -join ' '
            $coverCount = [regex]::Match($coverText,'(?i)Number\s+of\s+Records\s*:\s*(?<count>[\d,]+)')
            if ($coverCount.Success) { $expectedCount = [int]$coverCount.Groups['count'].Value.Replace(',','') }
        }
        $engine = [Tabula.Extractors.SpreadsheetExtractionAlgorithm]::new()
        for ($pageNumber=$StartPage; $pageNumber -le $document.NumberOfPages; $pageNumber++) {
            if ($ProgressCallback) { & $ProgressCallback $pageNumber $document.NumberOfPages $records.Count 'ParsingPage' | Out-Null }
            $page = $document.GetPage($pageNumber)
            if (!$page.Letters.Count) { throw "第 $pageNumber 頁無可讀取文字（可能是掃描頁）；首版不執行 OCR。" }
            if ([Math]::Abs([double]$page.Rotation.Value) % 360 -ne 0) { throw "第 $pageNumber 頁含旋轉，無法安全還原此版型，請先將頁面轉正。" }
            $area = [Tabula.ObjectExtractor]::Extract($document,$pageNumber)
            $tables = @($engine.Extract($area) | Sort-Object Top -Descending)
            if (!$tables.Count) { throw "第 $pageNumber 頁找不到可還原的格線表格。" }
            $header = Get-ReportHeader $page $tables[0].Top
            $pageRow = 0
            foreach ($table in $tables) {
                if ($table.ColumnCount -lt 2 -or $table.RowCount -lt 1) { throw "第 $pageNumber 頁的表格結構不明確。" }
                $tableText=[PdfReportExtractor.TableTextReader]::Read($page.Letters,$table)
                $columns = $null
                $firstDataRow = 1
                $firstCells = @($table.Rows[0])
                $firstValues = @($tableText[0] | ForEach-Object { ([regex]::Replace($_,'\s+',' ')).Trim() })
                $previous = $knownHeaders[$header.Title]
                if ($previous -and ($firstValues -join [char]31) -cne ($previous.Columns -join [char]31)) {
                    # Some report pages omit the repeated header. Reuse a proven schema
                    # only when every column boundary matches and the first row is taller
                    # than that schema's header. Otherwise reject an ambiguous layout.
                    $sameGrid = $firstCells.Count -eq $previous.Cells.Count
                    if ($sameGrid) {
                        for ($c=0; $c -lt $firstCells.Count; $c++) {
                            if ([Math]::Abs($firstCells[$c].Left-$previous.Cells[$c].Left) -gt 0.1 -or [Math]::Abs($firstCells[$c].Right-$previous.Cells[$c].Right) -gt 0.1) { $sameGrid=$false; break }
                        }
                    }
                    $firstHeight = ($firstCells | Measure-Object Height -Maximum).Maximum
                    if ($sameGrid -and $firstHeight -gt $previous.Height+0.5) {
                        $columns = $previous.Columns
                        $firstDataRow = 0
                    } elseif ($sameGrid -and [Math]::Abs($firstHeight-$previous.Height) -gt 0.5) {
                        throw "第 $pageNumber 頁首列與已知表頭不符，且無法確認延續表格，未寫入此檔。"
                    }
                }
                $rowIndex = 0
                foreach ($row in $table.Rows) {
                    $cells = @($row)
                    if ($cells.Count -ne $table.ColumnCount -or @($cells | Where-Object { $_.IsPlaceholder -or $_.IsSpanning -or $_.Width -le 0 -or $_.Height -le 0 }).Count) {
                        throw "第 $pageNumber 頁含不規則或合併儲存格，無法安全對應欄位。"
                    }
                    $values = @($tableText[$rowIndex])
                    if ($rowIndex -lt $firstDataRow) {
                        $columns = @($values | ForEach-Object { ([regex]::Replace($_,'\s+',' ')).Trim() })
                        if (@($columns | Where-Object { !$_ -or $_ -like '__*' }).Count -or @($columns | Sort-Object -Unique).Count -ne $columns.Count) {
                            throw "第 $pageNumber 頁表頭有空白、重名或保留欄名，無法安全取值。"
                        }
                        $schemas.Add([pscustomobject]@{Title=$header.Title;Columns=$columns})
                        $knownHeaders[$header.Title] = [pscustomobject]@{Columns=$columns;Cells=$cells;Height=($cells | Measure-Object Height -Maximum).Maximum}
                    } elseif (@($values | Where-Object { $_ }).Count) {
                        $normalized = @($values | ForEach-Object { ([regex]::Replace($_,'\s+',' ')).Trim() })
                        if (($normalized -join [char]31) -eq ($columns -join [char]31)) { $rowIndex++; continue }
                        if (@($values | Where-Object { $_.Length -gt 32767 }).Count) { throw "第 $pageNumber 頁有儲存格超過 Excel 的 32767 字元限制。" }
                        $pageRow++
                        $record = [ordered]@{
                            __ReportType=$header.Title; __SourceFile=[IO.Path]::GetFileName($Path); __SourcePath=$Path
                            __SourcePage=[string]$pageNumber; __SourceRow=[string]$pageRow; __ReportRange=$header.Range
                            __ImportedAt=$ImportedAt; __BatchId=$BatchId; __FileHash=$hash
                        }
                        for ($c=0; $c -lt $columns.Count; $c++) { $record[$columns[$c]] = $values[$c] }
                        $records.Add([pscustomobject]$record)
                    }
                    $rowIndex++
                }
            }
            if ($ProgressCallback) { & $ProgressCallback $pageNumber $document.NumberOfPages $records.Count 'PageComplete' | Out-Null }
        }
        if ($null -ne $expectedCount -and $records.Count -ne $expectedCount) {
            throw "萃取 $($records.Count) 筆，與封面 Number of Records 的 $expectedCount 筆不符，未寫入此檔。"
        }
        return [pscustomobject]@{Records=$records.ToArray();Schemas=$schemas.ToArray();Hash=$hash}
    } finally { $document.Dispose() }
}

function Get-WorksheetHeaders($Sheet) {
    if (!$Sheet.Dimension) { return @() }
    return @((1..$Sheet.Dimension.End.Column) | ForEach-Object { [string]$Sheet.Cells[1,$_].Value })
}

function Get-ReportSheetName([string]$Title, $Package) {
    $safe = [regex]::Replace($Title,'[\x00-\x1f\[\]:*?"<>|/\\]','_').Trim().Trim([char]39)
    if (!$safe) { $safe = 'Report' }
    if ($safe.Length -le 31 -and $safe -notin @('_Reports','_Imports') -and !$Package.Workbook.Worksheets[$safe]) { return $safe }
    $hash = (Get-TextHash $Title).Substring(0,8)
    $stem = $safe.Substring(0,[Math]::Min(22,$safe.Length))
    $candidate = $stem + '_' + $hash
    $counter = 1
    while ($Package.Workbook.Worksheets[$candidate]) {
        $suffix = '_' + $hash + '_' + $counter++
        $candidate = $safe.Substring(0,[Math]::Min(31-$suffix.Length,$safe.Length)) + $suffix
    }
    return $candidate
}

function Ensure-IndexSheet($Package, [string]$Name, [string[]]$Headers) {
    $sheet = $Package.Workbook.Worksheets[$Name]
    if (!$sheet) {
        $sheet = $Package.Workbook.Worksheets.Add($Name)
        for ($c=0;$c -lt $Headers.Count;$c++) { $sheet.Cells[1,($c+1)].Value=$Headers[$c] }
        $sheet.Row(1).Style.Font.Bold=$true
    } elseif ((@(Get-WorksheetHeaders $sheet) -join [char]31) -cne ($Headers -join [char]31)) { throw "歷史 Excel 的 $Name 索引格式不符，未寫入。" }
    return $sheet
}

function Add-ReportResults($Package, [object[]]$Results, [object[]]$Logs, [string]$ImportedAt) {
    $index = Ensure-IndexSheet $Package '_Reports' $script:IndexColumns
    $imports = Ensure-IndexSheet $Package '_Imports' $script:ImportColumns
    $mapping = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    if ($index.Dimension.End.Row -ge 2) {
        for ($r=2;$r -le $index.Dimension.End.Row;$r++) {
            $title=[string]$index.Cells[$r,1].Value; $name=[string]$index.Cells[$r,2].Value
            if (!$title -or !$name -or !$Package.Workbook.Worksheets[$name] -or $mapping.ContainsKey($title)) { throw '歷史 Excel 的報告索引不完整或重複，未寫入。' }
            $mapping.Add($title,[pscustomobject]@{Name=$name;IndexRow=$r})
        }
    }
    # Migrate report details inside the staged workbook. PDF columns cannot use
    # reserved __ names, so removing these exact legacy fields preserves source data.
    foreach ($item in $mapping.Values) {
        $sheet=$Package.Workbook.Worksheets[$item.Name]
        $headers=@(Get-WorksheetHeaders $sheet)
        if (@($headers | Sort-Object -Unique).Count -ne $headers.Count -or @($script:MetaColumns | Where-Object { $_ -cnotin $headers }).Count) { throw "工作表 $($sheet.Name) 的欄位結構已變更，未寫入。" }
        for ($c=$headers.Count-1; $c -ge 0; $c--) {
            if ($headers[$c] -cin $script:LegacyMetaColumns) { $sheet.DeleteColumn(($c+1)) }
        }
        $remaining=@(Get-WorksheetHeaders $sheet)
        if (@($remaining | Where-Object { $_ -like '__*' -and $_ -cnotin $script:MetaColumns }).Count) { throw "工作表 $($sheet.Name) 含未知的保留欄位，未寫入。" }
    }
    foreach ($result in $Results) {
        foreach ($schema in $result.Schemas) {
            if (!$mapping.ContainsKey($schema.Title)) {
                $name=Get-ReportSheetName $schema.Title $Package
                $sheet=$Package.Workbook.Worksheets.Add($name)
                $indexRow=$index.Dimension.End.Row+1
                $index.Cells[$indexRow,1].Value=$schema.Title; $index.Cells[$indexRow,2].Value=$name
                $mapping.Add($schema.Title,[pscustomobject]@{Name=$name;IndexRow=$indexRow})
                $headers=@($script:MetaColumns)
                for ($c=0;$c -lt $headers.Count;$c++) { $sheet.Cells[1,($c+1)].Value=$headers[$c] }
                $sheet.Row(1).Style.Font.Bold=$true
                $sheet.View.FreezePanes(2,1)
            }
            $sheet=$Package.Workbook.Worksheets[$mapping[$schema.Title].Name]
            $headers=@(Get-WorksheetHeaders $sheet)
            if (@($headers | Sort-Object -Unique).Count -ne $headers.Count -or @($script:MetaColumns | Where-Object { $_ -notin $headers }).Count) { throw "工作表 $($sheet.Name) 的欄位結構已變更，未寫入。" }
            foreach ($column in $schema.Columns) {
                if ($column -cnotin $headers) {
                    if ($column -in $headers) { throw "工作表 $($sheet.Name) 有僅大小寫不同的欄位，請先確認。" }
                    $sheet.Cells[1,($headers.Count+1)].Value=$column; $headers+=@($column)
                }
            }
        }
        foreach ($record in $result.Records) {
            $sheet=$Package.Workbook.Worksheets[$mapping[$record.__ReportType].Name]
            $headers=@(Get-WorksheetHeaders $sheet)
            $r=$sheet.Dimension.End.Row+1
            if ($r -gt 1048576) { throw "工作表 $($sheet.Name) 已達 Excel 列數限制，未寫入。" }
            for ($c=0;$c -lt $headers.Count;$c++) { $sheet.Cells[$r,($c+1)].Value=[string]$record.PSObject.Properties[$headers[$c]].Value }
        }
    }
    foreach ($title in $mapping.Keys) {
        $item=$mapping[$title]; $sheet=$Package.Workbook.Worksheets[$item.Name]
        $index.Cells[$item.IndexRow,3].Value=[int]($sheet.Dimension.End.Row-1)
        $index.Cells[$item.IndexRow,4].Value=$ImportedAt
    }
    foreach ($log in $Logs) {
        $r=$imports.Dimension.End.Row+1
        for ($c=0;$c -lt $script:ImportColumns.Count;$c++) { $imports.Cells[$r,($c+1)].Value=[string]$log.PSObject.Properties[$script:ImportColumns[$c]].Value }
    }
}

Export-ModuleMember -Function Initialize-Extractor,Read-ReportFile,Add-ReportResults,Get-TextHash

# Developer-only build helper; the packaged importer loads the compiled DLL.
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$lib=Join-Path $root 'lib'
$references=@(Get-ChildItem -LiteralPath $lib -Filter '*.dll' | Where-Object Name -ne 'PdfReportExtractor.TableText.dll' | ForEach-Object FullName)
$temporary=Join-Path ([IO.Path]::GetTempPath()) ('table-text-'+[guid]::NewGuid().ToString('N')+'.dll')
try {
    Add-Type -Path (Join-Path $root 'src/TableTextReader.cs') -ReferencedAssemblies $references -OutputAssembly $temporary -OutputType Library
    Copy-Item -LiteralPath $temporary -Destination (Join-Path $lib 'PdfReportExtractor.TableText.dll') -Force
    Write-Output 'Built lib/PdfReportExtractor.TableText.dll'
} finally { if(Test-Path -LiteralPath $temporary) {Remove-Item -LiteralPath $temporary} }

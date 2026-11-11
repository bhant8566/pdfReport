# Developer-only packaging helper. The importer never downloads dependencies.
[CmdletBinding()]
param([string]$ImportExcelSource = '')
$ErrorActionPreference = 'Stop'
if (!$ImportExcelSource) {
    $module = Get-Module -ListAvailable ImportExcel | Where-Object Version -eq ([version]'7.8.10') | Select-Object -First 1
    if (!$module) { throw 'Packaging requires ImportExcel 7.8.10. Supply -ImportExcelSource with its module directory.' }
    $ImportExcelSource = $module.ModuleBase
}
$root = Split-Path $PSScriptRoot -Parent
$lib = Join-Path $root 'lib'
$licenses = Join-Path $root 'licenses'
$cache = Join-Path ([IO.Path]::GetTempPath()) ('pdf-extractor-packages-' + [guid]::NewGuid().ToString('N'))
foreach ($path in @($lib,$licenses,$cache)) { [IO.Directory]::CreateDirectory($path) | Out-Null }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$packages = [ordered]@{
    'Tabula' = '1.0.1'
    'PdfPig' = '0.1.14'
    'Microsoft.Bcl.HashCode' = '6.0.0'
    'System.Memory' = '4.6.0'
    'System.Buffers' = '4.6.0'
    'System.Numerics.Vectors' = '4.6.0'
    'System.Runtime.CompilerServices.Unsafe' = '6.1.0'
    'System.ValueTuple' = '4.5.0'
}
$manifest = @()
foreach ($name in $packages.Keys) {
    $version = $packages[$name]
    $id = $name.ToLowerInvariant()
    $url = "https://api.nuget.org/v3-flatcontainer/$id/$version/$id.$version.nupkg"
    $archive = Join-Path $cache "$name.$version.zip"
    Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $archive
    $unpacked = Join-Path $cache $name
    [IO.Compression.ZipFile]::ExtractToDirectory($archive,$unpacked)
    $framework = $null
    foreach ($candidate in @('net462','net461','net46','net45','netstandard2.0')) {
        $candidatePath = Join-Path $unpacked "lib/$candidate"
        if (Test-Path -LiteralPath $candidatePath) { $framework = $candidate; break }
    }
    if (!$framework) { throw "No compatible assemblies found: $name" }
    Get-ChildItem -LiteralPath (Join-Path $unpacked "lib/$framework") -Filter '*.dll' | Copy-Item -Destination $lib -Force
    $packageLicense = Join-Path $licenses "$name-$version"
    [IO.Directory]::CreateDirectory($packageLicense) | Out-Null
    Get-ChildItem -LiteralPath $unpacked -File | Where-Object { $_.Name -match 'license|notice|\.nuspec$' } | Copy-Item -Destination $packageLicense -Force
    if (!(Get-ChildItem -LiteralPath $packageLicense -File | Where-Object Name -match 'license')) {
        [xml]$spec = Get-Content -LiteralPath (Get-ChildItem -LiteralPath $unpacked -Filter '*.nuspec' | Select-Object -First 1).FullName -Raw
        $expression = [string]$spec.package.metadata.license.'#text'
        if (!$expression) { throw "Missing package license: $name" }
        Invoke-WebRequest -UseBasicParsing -Uri ('https://licenses.nuget.org/' + [uri]::EscapeDataString($expression)) -OutFile (Join-Path $packageLicense 'LICENSE.html')
    }
    $manifest += [pscustomobject]@{package=$name;version=$version;framework=$framework;source=$url;sha256=(Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash}
    Write-Output "Packaged $name $version ($framework)"
}
$moduleTarget = Join-Path $lib 'ImportExcel'
[IO.Directory]::CreateDirectory($moduleTarget) | Out-Null
foreach ($name in @('ImportExcel.psd1','ImportExcel.psm1','EPPlus.dll','Plot.ps1','LICENSE.txt','Public','Private','Charting','InferData','Pivot','en')) {
    Copy-Item -LiteralPath (Join-Path $ImportExcelSource $name) -Destination $moduleTarget -Recurse -Force
}
Copy-Item -LiteralPath (Join-Path $ImportExcelSource 'LICENSE.txt') -Destination (Join-Path $licenses 'ImportExcel-LICENSE.txt') -Force
Invoke-WebRequest -UseBasicParsing -Uri 'https://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt' -OutFile (Join-Path $licenses 'EPPlus-LGPL-2.1.txt')
$manifest += [pscustomobject]@{package='ImportExcel';version='7.8.10';source='https://github.com/dfinke/ImportExcel';sha256=(Get-FileHash (Join-Path $moduleTarget 'EPPlus.dll') -Algorithm SHA256).Hash}
& (Join-Path $PSScriptRoot 'Build-TextReader.ps1')
$manifest += [pscustomobject]@{package='PdfReportExtractor.TableText';version='1.0.0';source='src/TableTextReader.cs';sha256=(Get-FileHash (Join-Path $lib 'PdfReportExtractor.TableText.dll') -Algorithm SHA256).Hash}
[IO.File]::WriteAllText((Join-Path $root 'dependencies.json'),(ConvertTo-Json -InputObject $manifest -Depth 5),[Text.UTF8Encoding]::new($false))
Write-Output "Package cache: $cache"

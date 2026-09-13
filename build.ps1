$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
dotnet build src/Obj2Pdf/Obj2Pdf.csproj -c Release --nologo
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
dotnet build src/Crop/Crop.csproj -c Release --nologo
if ($LASTEXITCODE -ne 0) { throw 'Crop build failed' }
Get-ChildItem src/Crop/bin/Release/net48 -File | Where-Object { $_.Extension -in '.dll','.exe','.config' } | Copy-Item -Destination src/Obj2Pdf/bin/Release/net48 -Force
New-Item -ItemType Directory -Force artifacts, dist | Out-Null
$payload = Join-Path $PSScriptRoot 'artifacts/payload.zip'
if (Test-Path -LiteralPath $payload) { Remove-Item -LiteralPath $payload }
$files = @(Get-ChildItem src/Crop/bin/Release/net48 -File | Where-Object { $_.Extension -in '.dll','.exe','.config' } | ForEach-Object FullName)
$files += (Join-Path $PSScriptRoot 'src/Obj2Pdf/bin/Release/net48/Obj2Pdf.dll')
$files += (Join-Path $PSScriptRoot 'THIRD-PARTY-NOTICES.txt')
Compress-Archive -LiteralPath $files -DestinationPath $payload
$compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
if (!(Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/csc.exe' }
& $compiler /nologo /target:winexe /platform:anycpu /optimize+ '/out:dist\obj2pdf-Setup.exe' /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll '/resource:artifacts\payload.zip,payload.zip' 'src\Setup\Program.cs'
if ($LASTEXITCODE -ne 0) { throw 'Setup compilation failed' }
Get-FileHash dist/obj2pdf-Setup.exe -Algorithm SHA256 | Format-List

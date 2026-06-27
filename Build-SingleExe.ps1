Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$solution = (Get-ChildItem -Path "." -Filter "*.sln" | Select-Object -First 1).FullName
$outputDir = ".\bin\Release"
$exePath = Join-Path $outputDir "AutomatizationOfSW.exe"

dotnet clean $solution -c Release
if ($LASTEXITCODE -ne 0) {
    throw "dotnet clean failed"
}

dotnet build $solution -c Release
if ($LASTEXITCODE -ne 0) {
    throw "dotnet build failed"
}

$dlls = @(Get-ChildItem -Path $outputDir -Filter "*.dll" -Recurse -ErrorAction SilentlyContinue)
if ($dlls.Count -gt 0) {
    throw "Single-exe build produced DLL files in $outputDir"
}

if (!(Test-Path $exePath)) {
    throw "Build succeeded but $exePath was not found"
}

& $exePath --self-test

Write-Host "Single-exe build is ready:"
Write-Host (Resolve-Path $exePath)

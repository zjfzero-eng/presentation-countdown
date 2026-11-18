$ErrorActionPreference = 'Stop'
$frameworkRoot = Join-Path $env:WINDIR 'Microsoft.NET'
$compiler = Join-Path $frameworkRoot 'Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) {
    $compiler = Join-Path $frameworkRoot 'Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path -LiteralPath $compiler)) {
    throw 'The .NET Framework C# compiler was not found.'
}

$source = Join-Path $PSScriptRoot 'src\Countdown.cs'
$icon = Join-Path $PSScriptRoot 'src\countdown.ico'
$output = Join-Path $PSScriptRoot 'dist\presentation-countdown.exe'
New-Item -ItemType Directory -Path (Split-Path -Parent $output) -Force | Out-Null
& $compiler /nologo /target:winexe /optimize+ /codepage:65001 "/win32icon:$icon" "/out:$output" $source
if ($LASTEXITCODE -ne 0) {
    throw "Compilation failed with exit code $LASTEXITCODE."
}
Get-Item -LiteralPath $output

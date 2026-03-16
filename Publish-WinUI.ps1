param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSCommandPath
$msbuild = 'C:\Program Files\Microsoft Visual Studio\18\Insiders\MSBuild\Current\Bin\amd64\MSBuild.exe'
$winuiProject = Join-Path $root 'PptxVideoProcessing.WinUI\PptxVideoProcessing.WinUI.csproj'
$workerProject = Join-Path $root 'PptxVideoProcessing\PptxVideoProcessing.vcxproj'
$publishDir = Join-Path $root ("Publish\win-x64\$Configuration")

if (-not (Test-Path $msbuild)) {
    throw "未找到 MSBuild：$msbuild"
}

& $msbuild $workerProject /p:Configuration=$Configuration /p:Platform=x64 /m
if ($LASTEXITCODE -ne 0) {
    throw 'native worker 构建失败。'
}

$env:DOTNET_CLI_HOME = Join-Path $root '.dotnet-cli-home'
if (-not (Test-Path $env:DOTNET_CLI_HOME)) {
    New-Item -ItemType Directory -Path $env:DOTNET_CLI_HOME | Out-Null
}

if (Test-Path $publishDir) {
    Remove-Item -Recurse -Force $publishDir
}

$publishArgs = @(
    'publish',
    $winuiProject,
    '-c', $Configuration,
    '-r', 'win-x64',
    '-p:Platform=x64',
    '-p:BuildingSolutionFile=true',
    '-p:PublishSingleFile=true',
    '-p:SelfContained=true',
    '-p:WindowsAppSDKSelfContained=true',
    '-p:EnableMsixTooling=true',
    '-p:DebugType=None',
    '-p:DebugSymbols=false',
    '-o', $publishDir
)

& dotnet @publishArgs
if ($LASTEXITCODE -ne 0) {
    throw 'WinUI 发布失败。'
}

Write-Host "发布完成：$publishDir"


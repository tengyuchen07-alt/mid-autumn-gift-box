$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$testProject = Join-Path $projectRoot 'tests\MidAutumnGiftBox.Core.Tests\MidAutumnGiftBox.Core.Tests.csproj'
$appProject = Join-Path $projectRoot 'src\MidAutumnGiftBox.App\MidAutumnGiftBox.App.csproj'
$publishDirectory = Join-Path $projectRoot 'build\publish'

dotnet restore $testProject --nologo
if ($LASTEXITCODE -ne 0) { throw "Test project restore failed with exit code $LASTEXITCODE." }

dotnet restore $appProject --runtime win-x64 --nologo
if ($LASTEXITCODE -ne 0) { throw "App project restore failed with exit code $LASTEXITCODE." }

dotnet run --project $testProject --no-restore
if ($LASTEXITCODE -ne 0) { throw "Core tests failed with exit code $LASTEXITCODE." }

dotnet publish $appProject --configuration Release --runtime win-x64 --self-contained true --no-restore --nologo `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    --output $publishDirectory
if ($LASTEXITCODE -ne 0) { throw "App publish failed with exit code $LASTEXITCODE." }

Write-Host "Published to $publishDirectory"

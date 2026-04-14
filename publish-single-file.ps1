$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$env:DOTNET_CLI_HOME = Join-Path $repoRoot ".dotnet"
$env:NUGET_PACKAGES = Join-Path $repoRoot ".nuget"

dotnet publish (Join-Path $repoRoot "TaskManager\TaskManager.csproj") `
  -p:PublishProfile=Properties\PublishProfiles\SingleFile-win-x64.pubxml

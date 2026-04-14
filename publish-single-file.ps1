$ErrorActionPreference = "Stop"

$env:DOTNET_CLI_HOME = "c:\Work\Project-netstat\.dotnet"
$env:NUGET_PACKAGES = "c:\Work\Project-netstat\.nuget"

dotnet publish "TaskManager\TaskManager.csproj" `
  -p:PublishProfile=Properties\PublishProfiles\SingleFile-win-x64.pubxml

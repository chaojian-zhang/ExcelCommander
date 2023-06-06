$PrevPath = Get-Location

Write-Host "Publish for Final Packaging build."
Set-Location $PSScriptRoot

$PublishFolder = "$PSScriptRoot\..\Publish"
$LibraryPublishFolder = "$PublishFolder\Libraries"
$NugetPublishFolder = "$PublishFolder\Nugets"

# Delete current data
Remove-Item $PublishFolder -Recurse -Force

# Publish Executables
$PublishExecutables = @(
    "ExcelCommander"
)
foreach ($Item in $PublishExecutables)
{
    dotnet publish $PSScriptRoot\..\$Item\$Item.csproj --use-current-runtime --output $PublishFolder
}
# Publish Windows-only Executables
$PublishWindowsExecutables = @(
)
foreach ($Item in $PublishWindowsExecutables)
{
    dotnet publish $PSScriptRoot\..\$Item\$Item.csproj --runtime win-x64 --self-contained --output $PublishFolder
}
# Publish Loose Libraries
$PublishLibraries = @(
    "ExcelCommander.Base"
)
foreach ($Item in $PublishLibraries)
{
    dotnet publish $PSScriptRoot\..\$Item\$Item.csproj --use-current-runtime --output $LibraryPublishFolder
}
# Publish Nugets
$PublishNugets = @(
    "ExcelCommander"
	"ExcelCommander.Base"
)
foreach ($Item in $PublishNugets)
{
    dotnet pack $PSScriptRoot\..\$Item\$Item.csproj --output $NugetPublishFolder
}

# Create archive
$Date = Get-Date -Format yyyyMMdd
$ArchiveFolder = "$PublishFolder\..\Packages"
$ArchivePath = "$ArchiveFolder\ExcelCommander_DistributionBuild_Windows_B$Date.zip"
New-Item -ItemType Directory -Force -Path $ArchiveFolder
Compress-Archive -Path $PublishFolder\* -DestinationPath $ArchivePath -Force

# Validation
if (-Not (Test-Path (Join-Path $PublishFolder "ExcelCommander.exe")))
{
    Write-Host "Build failed."
    Exit
}

# Notes
Write-Host "Excel Addins and XlsxCommander must be published manually (or likely with .Net Framework build tools?)"

Set-Location $PrevPath
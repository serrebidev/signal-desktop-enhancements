$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestPath = Join-Path $repoRoot "manifest.ini"

$versionLine = Select-String -Path $manifestPath -Pattern '^version = "(.+)"$'
if (-not $versionLine) {
	throw "Could not determine version from manifest.ini"
}

$version = $versionLine.Matches[0].Groups[1].Value
$distDir = Join-Path $repoRoot "dist"
$stageDir = Join-Path $distDir "stage"
$zipPath = Join-Path $distDir ("signalDesktopEnhancements-{0}.zip" -f $version)
$addonPath = Join-Path $distDir ("signalDesktopEnhancements-{0}.nvda-addon" -f $version)

if (Test-Path $stageDir) {
	Remove-Item -LiteralPath $stageDir -Recurse -Force
}
if (Test-Path $zipPath) {
	Remove-Item -LiteralPath $zipPath -Force
}
if (Test-Path $addonPath) {
	Remove-Item -LiteralPath $addonPath -Force
}

New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repoRoot "manifest.ini") -Destination $stageDir
Copy-Item -LiteralPath (Join-Path $repoRoot "README.md") -Destination $stageDir
Copy-Item -LiteralPath (Join-Path $repoRoot "appModules") -Destination $stageDir -Recurse

$pycache = Join-Path $stageDir "appModules\\__pycache__"
if (Test-Path $pycache) {
	Remove-Item -LiteralPath $pycache -Recurse -Force
}

Compress-Archive -Path (Join-Path $stageDir '*') -DestinationPath $zipPath -Force
Move-Item -LiteralPath $zipPath -Destination $addonPath -Force
Remove-Item -LiteralPath $stageDir -Recurse -Force

Write-Output ("Built {0}" -f $addonPath)

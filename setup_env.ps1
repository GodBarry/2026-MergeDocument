param(
    [switch]$FullInstall,
    [switch]$SkipOCR,
    [switch]$SkipAI
)

$script = Join-Path $PSScriptRoot 'setup_env.py'
$setupArgs = @()
if ($FullInstall) { $setupArgs += '--full-install' }
if ($SkipOCR) { $setupArgs += '--skip-ocr' }
if ($SkipAI) { $setupArgs += '--skip-ai' }

python $script @setupArgs
exit $LASTEXITCODE

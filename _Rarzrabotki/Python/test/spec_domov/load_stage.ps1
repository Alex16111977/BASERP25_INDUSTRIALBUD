# Load one batch of P1 into bas_industrialbud: gates -> db-load-xml Partial -> log -> /CheckModules -> db-update -Dynamic + -> gate.
# Loads ONLY from the main dump directory (D1). ASCII-only file: Cyrillic checks live in Python (check_log.py).
param(
    [Parameter(Mandatory=$true)][string]$List,
    [Parameter(Mandatory=$true)][string]$Tag,
    [switch]$WithRoles,
    [switch]$DryRun
)
$V8   = "C:\Program Files\1cv8\8.3.20.1914\bin"
$SK   = "C:\Configuration_downloads\BASERP25\.claude\skills"
$CFG  = "C:\Configuration_downloads\BASERP25\_Rarzrabotki\BASEBuh"
$HERE = Split-Path -Parent $MyInvocation.MyCommand.Path
$LOGD = "C:\Users\SUPPOR~1\AppData\Local\Temp\claude\plan_p1\load_$Tag"
New-Item -ItemType Directory -Force $LOGD | Out-Null
$CONN = @("-V8Path", $V8, "-InfoBaseServer", "localhost", "-InfoBaseRef", "bas_industrialbud", "-UserName", "cfo", "-Password", "2442")
# Summary of a real load (committed): out\<Tag>_load_<yyyyMMdd_HHmm>.txt. Dry run writes nothing.
$SUMMARY = $null
if (-not $DryRun) {
    New-Item -ItemType Directory -Force (Join-Path $HERE "out") | Out-Null
    $SUMMARY = Join-Path (Join-Path $HERE "out") ("{0}_load_{1}.txt" -f $Tag, (Get-Date -Format "yyyyMMdd_HHmm"))
}

function Note([string]$Text) {
    Write-Host $Text
    if ($SUMMARY) { Add-Content -Path $SUMMARY -Value $Text -Encoding UTF8 }
}

function Invoke-Step([string]$Name, [scriptblock]$Body) {
    Note "=== $Name"
    & $Body
    $code = $LASTEXITCODE
    Note "    exit=$code"
    if ($code -ne 0) { Note "STOP at step '$Name' exit=$code. Logs: $LOGD"; exit 1 }
}

Note ("load {0}: list {1}, started {2}" -f $Tag, $List, (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))

Invoke-Step "sessions"        { python -X utf8 "$HERE\sessions_gate.py" }
Invoke-Step "main = worktree" { python -X utf8 "$HERE\to_main.py" --verify $List }
Invoke-Step "preflight config" { python -X utf8 "$HERE\preflight_config.py" }
if ($WithRoles) { Invoke-Step "preflight roles" { python -X utf8 "$HERE\preflight_roles.py" } }
if ($DryRun) { Write-Host "DRY RUN OK ($Tag): gates passed, nothing loaded"; exit 0 }

$files = (Get-Content $List -Encoding UTF8 | Where-Object { $_.Trim() -and -not $_.StartsWith("#") }) -join ","
Invoke-Step "db-load-xml Partial" {
    & powershell.exe -NoProfile -File "$SK\db-load-xml\scripts\db-load-xml.ps1" @CONN -ConfigDir $CFG -Mode Partial -Files $files -StrictLog *>&1 | Tee-Object -FilePath "$LOGD\load.log"
}
Invoke-Step "load log" { python -X utf8 "$HERE\check_log.py" load "$LOGD\load.log" }
Invoke-Step "CheckModules" {
    $log = "$LOGD\checkmodules.log"
    $p = Start-Process -FilePath "$V8\1cv8.exe" -ArgumentList @("DESIGNER", "/S", "`"localhost/bas_industrialbud`"", "/N`"cfo`"", "/P`"2442`"", "/CheckModules", "-ThinClient", "-Server", "-ExternalConnection", "/Out", "`"$log`"", "/DisableStartupDialogs") -NoNewWindow -Wait -PassThru
    python -X utf8 "$HERE\check_log.py" checkmodules $log $p.ExitCode
}
Invoke-Step "db-update -Dynamic +" {
    & powershell.exe -NoProfile -File "$SK\db-update\scripts\db-update.ps1" @CONN -Dynamic "+" *>&1 | Tee-Object -FilePath "$LOGD\update.log"
}
Invoke-Step "preflight config after" { python -X utf8 "$HERE\preflight_config.py" }
Note "OK load $Tag. Logs: $LOGD"
exit 0

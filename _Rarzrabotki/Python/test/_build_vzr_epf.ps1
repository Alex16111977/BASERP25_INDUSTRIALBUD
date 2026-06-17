# Автономная сборка СинхронизироватьВзаиморасчеты.epf через временную stub-базу 8.3.20
# (патчим СГЕНЕРИРОВАННЫЙ артефакт stub-cfg, НЕ скил). Не трогает рабочие базы.
$ErrorActionPreference = 'Stop'
$plat = 'C:\Program Files\1cv8\8.3.20.1914\bin\1cv8.exe'
$stub = "$env:TEMP\epf_stub_db_1225142637\cfg"
$work = "$env:TEMP\vzr_build_work"
$ib   = "$env:TEMP\vzr_build_ib"
$src  = 'C:\Configuration_downloads\BASERP25\.claude\worktrees\silly-colden-130551\_Rarzrabotki\Обработки\СинхронизироватьВзаиморасчеты.xml'
$out  = 'C:\Configuration_downloads\BASERP25\.claude\worktrees\silly-colden-130551\_Rarzrabotki\Обработки\СинхронизироватьВзаиморасчеты.epf'
$log  = "$env:TEMP\vzr_build_log.txt"

if (Test-Path $work) { Remove-Item $work -Recurse -Force }
if (Test-Path $ib)   { Remove-Item $ib -Recurse -Force }
New-Item -ItemType Directory -Path $work -Force | Out-Null
Copy-Item "$stub\*" $work -Recurse -Force

Get-ChildItem $work -Recurse -Filter *.xml | ForEach-Object {
    $c = Get-Content $_.FullName -Raw
    $c = $c -replace 'version="2\.17"','version="2.13"' -replace 'Version8_3_24','Version8_3_20'
    [System.IO.File]::WriteAllText($_.FullName, $c, (New-Object System.Text.UTF8Encoding $true))
}
Write-Output ("patched xml: " + (Get-ChildItem $work -Recurse -Filter *.xml).Count + " файлов")

# Убрать типы, которые EPF НЕ использует как типы (стаб генерит битый Хозрасчетный + зависимости;
# Хозрасчетный/субконто/перечисления присутствуют только в строках запросов, не в типах реквизитов)
Remove-Item "$work\ChartsOfAccounts","$work\ChartsOfCharacteristicTypes","$work\Enums" -Recurse -Force -ErrorAction SilentlyContinue
$cfgPath = "$work\Configuration.xml"
$c = Get-Content $cfgPath -Raw
$c = $c -replace '(?m)^\s*<ChartOfAccounts>.*?</ChartOfAccounts>\s*\r?\n',''
$c = $c -replace '(?m)^\s*<Enum>.*?</Enum>\s*\r?\n',''
$c = $c -replace '(?m)^\s*<ChartOfCharacteristicTypes>.*?</ChartOfCharacteristicTypes>\s*\r?\n',''
[System.IO.File]::WriteAllText($cfgPath, $c, (New-Object System.Text.UTF8Encoding $true))
Write-Output "stripped ChartsOfAccounts/CCT/Enums from stub"

$csIB = "File=$ib;"

function Run1C([string[]]$a) {
    if (Test-Path $log) { Remove-Item $log -Force }
    $p = Start-Process -FilePath $plat -ArgumentList $a -Wait -PassThru -NoNewWindow
    $txt = ''
    if (Test-Path $log) { $txt = Get-Content $log -Raw }
    return @{ code = $p.ExitCode; log = $txt }
}

Write-Output '=== CREATEINFOBASE ==='
$r = Run1C @('CREATEINFOBASE', $csIB, '/DisableStartupDialogs')
Write-Output ("code=" + $r.code + " " + $r.log)
if ($r.code -ne 0) { throw 'createinfobase failed' }

Write-Output '=== LoadConfigFromFiles ==='
$r = Run1C @('DESIGNER', '/IBConnectionString', $csIB, '/LoadConfigFromFiles', $work, '/Out', $log, '/DisableStartupDialogs')
Write-Output ("code=" + $r.code)
Write-Output $r.log
if ($r.code -ne 0) { throw 'loadconfig failed' }

Write-Output '=== UpdateDBCfg ==='
$r = Run1C @('DESIGNER', '/IBConnectionString', $csIB, '/UpdateDBCfg', '/Out', $log, '/DisableStartupDialogs')
Write-Output ("code=" + $r.code)
Write-Output $r.log

Write-Output '=== Build EPF ==='
$r = Run1C @('DESIGNER', '/IBConnectionString', $csIB, '/LoadExternalDataProcessorOrReportFromFiles', $src, $out, '/Out', $log, '/DisableStartupDialogs')
Write-Output ("code=" + $r.code)
Write-Output $r.log
if ((Test-Path $out) -and ((Get-Item $out).Length -gt 0)) {
    Write-Output ("OK EPF built: " + (Get-Item $out).Length + " bytes, mtime " + (Get-Item $out).LastWriteTime)
} else {
    throw 'EPF not produced'
}

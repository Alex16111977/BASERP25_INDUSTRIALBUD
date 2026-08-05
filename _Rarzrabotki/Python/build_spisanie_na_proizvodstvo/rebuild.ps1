# Пересборка отчёта А_ОтчетПоСписаниюНаПроизводство: валидация СКД -> .erf -> копия в основной каталог
# Сборка против СВОБОДНОЙ базы BaseERPRazr — конфигуратор рабочей BaseERP занят (CLAUDE.md).
$ErrorActionPreference = "Stop"
$wt   = "C:\Configuration_downloads\BASERP25\.claude\worktrees\plan-fact-production-writeoff-6e8c2a\_Rarzrabotki\Отчеты"
$main = "C:\Configuration_downloads\BASERP25\_Rarzrabotki\Отчеты"
$name = "А_ОтчетПоСписаниюНаПроизводство"
$tpl  = "$wt\$name\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml"

Write-Host "--- 1. skd-validate ---"
powershell.exe -NoProfile -File "C:\Configuration_downloads\BASERP25\.claude\skills\skd-validate\scripts\skd-validate.ps1" -TemplatePath $tpl
if ($LASTEXITCODE -ne 0) { throw "skd-validate failed" }

Write-Host "--- 2. erf-build (BaseERPRazr) ---"
powershell.exe -NoProfile -File "C:\Configuration_downloads\BASERP25\.claude\skills\epf-build\scripts\epf-build.ps1" -InfoBaseServer "SQLSERVER" -InfoBaseRef "BaseERPRazr" -UserName "Администратор" -Password "24043" -SourceFile "$wt\$name.xml" -OutputFile "$wt\$name.erf"
if ($LASTEXITCODE -ne 0) { throw "erf-build failed" }

Write-Host "--- 3. копия в основной каталог конфигурации (Rule #4) ---"
Remove-Item -Recurse -Force "$main\$name" -ErrorAction SilentlyContinue
Copy-Item -Recurse "$wt\$name" "$main\$name"
Copy-Item "$wt\$name.xml" "$main\$name.xml"
Copy-Item "$wt\$name.erf" "$main\$name.erf"
Write-Host "DONE"

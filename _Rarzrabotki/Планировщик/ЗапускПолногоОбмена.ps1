# Автозапуск полного обмена по всем базам (планировщик заданий Windows).
# Открывает BaseERP толстым клиентом; обработка А_АвтоматическийОбменДанными по
# ПараметрЗапуска "SYNCALL" выполняет ЗапуститьОбменПоВсемБазам() и закрывается.
# ----- НАСТРОЙКА (правьте под свой сервер) -----
$V8  = 'C:\Program Files\1cv8\8.3.21.1302\bin\1cv8.exe'
$Epf = 'C:\!OBMIN-OLD-NEW\Обработки\А_АвтоматическийОбменДанными.epf'
$Srv = 'SQLSERVER\BaseERP'          # на текущей машине: 'localhost\BaseERP'
$Usr = 'Администратор'
$Pwd = '24043'
# ------------------------------------------------
$argList = @(
    'ENTERPRISE',
    '/S', $Srv,
    '/N', $Usr,
    '/P', $Pwd,
    '/Execute', $Epf,
    '/CSYNCALL',
    '/DisableStartupDialogs',
    '/DisableStartupMessages'
)
$proc = Start-Process -FilePath $V8 -ArgumentList $argList -Wait -PassThru
exit $proc.ExitCode

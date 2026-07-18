@echo off



set "V8=C:\Program Files\1cv8\8.3.20.1914\bin\1cv8.exe"
set "EPF=C:\!OBMIN-OLD-NEW\Obmen\Obmen.epf"
set "SRV=SQLSERVER\BaseERP"
set "USR=cfo"
set "PWD=2442"

"%V8%" ENTERPRISE /S "%SRV%" /N "%USR%" /P "%PWD%" /Execute "%EPF%" /CSYNCALL /DisableStartupDialogs /DisableStartupMessages

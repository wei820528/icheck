chcp 65001

@echo off
::網址
set HttpAll=https://localhost:447/iCheckAPI 
::時間
set HOUR=%time:~0,2% 
:: 今天結束時間 
set MaxDate=18
:: 製作表單時間
set TaskMax=8 
:: 稽催MAIL小於多少做動作
set NewSendMail=18 
:: 匯出CSV小於多少做動作
set SendSystem=18
c:
cd C:\Program Files (x86)\Internet Explorer
 if %HOUR% LEQ  %MaxDate% (
     if %HOUR% EQU  %TaskMax% (
        iexplore.exe  https://localhost:447/iCheckAPI/Task/Task
		ping -n 10 127.1 >nul 3>nul
     ) 
     if %HOUR% LEQ  %NewSendMail% (
         iexplore.exe  https://localhost:447/iCheckAPI/Task/SendSystem
		ping -n 10 127.1 >nul 3>nul 
	) 
	 if %HOUR% LEQ  %SendSystem% (
		iexplore.exe  https://localhost:447/iCheckAPI/Task/NewSendMail
		ping -n 10 127.1 >nul 3>nul
	 ) 
 ) 
taskkill /im iexplore.exe /f
exit



 
REM :Task_C
REM echo 產生表單
REM msedge.exe  %HttpAll%/Task/Task
REM @pause >nul
REM exit
 
REM :SendSystem_C
REM echo 匯出CSV
REM msedge.exe  %HttpAll%/Task/SendSystem  REM 匯出CSV
REM @pause >nul
REM exit

REM :SendMail_C
REM echo 稽催MAIL
REM msedge.exe  %HttpAll%/Task/NewSendMail REM 稽催MAIL
REM @pause >nul
REM exit
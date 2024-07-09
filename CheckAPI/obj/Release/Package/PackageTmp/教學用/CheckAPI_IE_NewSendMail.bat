chcp 65001

@echo off
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
	 if %HOUR% LEQ  %SendSystem% (
		iexplore.exe  https://localhost:447/iCheckAPI/Task/NewSendMail
		ping -n 10 127.1 >nul 3>nul
	 ) 
) 
taskkill /im iexplore.exe /f
exit
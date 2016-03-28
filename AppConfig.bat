@echo off

set time=%TIME: =0%
set hr=%time:~0,2%

for /f %%i in ('powershell ^(get-date^).DayOfWeek') do set dow=%%i

PowerShell.exe -Windowstyle Hidden -ExecutionPolicy Bypass -File "C:\Users\checkin\Google Drive\Scripts\ClosingF1Dialogue.ps1" 

Rem ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Sunday~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

IF /I "%dow%" EQU "Sunday" (

	IF /I "%hr%" EQU "07" ( rem This is setup For sunday 8:00 Service
		START /WAIT /d "C:\Users\checkin\Google Drive\Scripts" AppLoad.exe 
		START /d "C:\Users\checkin\Google Drive\Scripts" %hr%SS.exe -s
		TIMEOUT /t 900 /NOBREAK 
		TASKKILL /f /im %hr%SS.exe 
		
	)
	
	IF /I "%hr%" EQU "09" ( rem This is setup for 10:00 Service
	
		START /WAIT /d "C:\Users\checkin\Google Drive\Scripts" AppLoad.exe 
		START /d "C:\Users\checkin\Google Drive\Scripts" %hr%SS.exe -s
		TIMEOUT /t 1800 /NOBREAK 
		TASKKILL /f /im %hr%SS.exe
	)
	
	IF /I "%hr%" EQU "10" ( rem This is setup for 11:30 Service
		START /WAIT /d "C:\Users\checkin\Google Drive\Scripts" AppLoad.exe 
		START /d "C:\Users\checkin\Google Drive\Scripts" %hr%SS.exe -s
		TIMEOUT /t 1800 /NOBREAK 
		TASKKILL /f /im %hr%SS.exe
	)

	IF /I "%hr%" EQU "12" ( rem This is the end of Sunday Services
		START /d "C:\Users\checkin\Google Drive\Scripts" %hr%SS.exe -s
		TIMEOUT /t 2700 /NOBREAK 
		START nircmd.exe monitor off
	)	
)
rem ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Wednesday~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

IF /I "%dow%" EQU "Wednesday" (
	IF /I "%hr%" EQU "17" ( rem this is Wednesday setup for life groups
		START /WAIT /d "C:\Users\checkin\Google Drive\Scripts" AppLoad.exe 
		
	)
	
	IF /I "%hr%" EQU "18" ( rem this is Wednesday setup
		START /WAIT /d "C:\Users\checkin\Google Drive\Scripts" AppLoad.exe 
		START /d "C:\Users\checkin\Google Drive\Scripts" %hr%SS.exe -s
		TIMEOUT /t 300 /NOBREAK 
		TASKKILL /f /im %hr%SS.exe
	)
	
	IF /I "%hr%" EQU "19" ( rem this is the end of Wednesday Service
		START /d "C:\Users\checkin\Google Drive\Scripts" %hr%SS.exe -s
		TIMEOUT /t 2700 /NOBREAK 
		START nircmd.exe monitor off
	)
)	

rem ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Tuesday~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

IF /I "%dow%" EQU "Tuesday" (
	IF /I "%hr%" EQU "17" ( rem this is Tuesday so Task will be different
		START /WAIT /d "C:\Users\checkin\Google Drive\Scripts" AppLoad.exe 
	)
	
	IF /I "%hr%" EQU "08" ( rem this is Tuesday so Task will be different
		START /WAIT /d "C:\Users\checkin\Google Drive\Scripts" AppLoad.exe 
	)
)

rem ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Friday~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

IF /I "%dow%" EQU "Friday" (
	IF /I "%hr%" EQU "18" ( rem this is Tuesday so Task will be different 
		START /WAIT /d "C:\Users\checkin\Google Drive\Scripts" AppLoad.exe 
		START /d "C:\Users\checkin\Google Drive\Scripts" %hr%SS.exe -s
		TIMEOUT /t 900 /NOBREAK 
		TASKKILL /f /im %hr%SS.exe
	)
)

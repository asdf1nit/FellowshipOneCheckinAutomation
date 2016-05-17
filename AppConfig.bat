@echo off

rem Time vars for selecting service
set time=%TIME: =0%
set hr=%time:~0,2%
rem Screen Saver vars for testing if  special exist
set SundaySpecial=%~dp0%SunSpecial.exe
set WednesdaySpecial=%~dp0%WedSpecial.exe
set LoftSpecial=%~dp0%LoftSpecial.exe
rem location and computer variables
set local=%~dp0
set comp=%COMPUTERNAME%
SET test=%comp:~0,4%
set loft=LOFT

rem sets the day of week
for /f %%i in ('powershell ^(get-date^).DayOfWeek') do set dow=%%i

rem Section for Special Screen savers
IF EXIST "%SundaySpecial%" (
set sunSS=SunSpecial.exe
) ELSE (
set sunSS=%hr%SS.exe
)

IF EXIST "%WednesdaySpecial%" (
set wedSS=WedSpecial.exe
) ELSE (
set wedSS=%hr%SS.exe
)

IF EXIST "%LoftSpecial%" (
set loftSS=LoftSpecial.exe
) ELSE (
set loftSS=%hr%SS.exe
)

PowerShell.exe -Windowstyle Hidden -ExecutionPolicy Bypass -File "%local%ClosingF1Dialogue.ps1" 

Rem ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Sunday~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

IF /I "%dow%" EQU "Sunday" (

	IF /I "%hr%" EQU "07" ( rem This is setup For sunday 8:00 Service
		START /WAIT /d "%local%" Appload.exe 
		
		IF /I %test% == %loft% (
		START /d "%local%" %loftSS% -s
		TIMEOUT /t 900 /NOBREAK 
		TASKKILL /f /im %loftSS%
		) ELSE (
		START /d "%local%" %sunSS% -s
		TIMEOUT /t 900 /NOBREAK 
		TASKKILL /f /im %sunSS%
		)
	)
	
	IF /I "%hr%" EQU "09" ( rem This is setup for 10:00 Service
		START /WAIT /d "%local%" Appload.exe 
		
		IF /I %test% == %loft% (
		START /d "%local%" %loftSS% -s
		TIMEOUT /t 1800 /NOBREAK 
		TASKKILL /f /im %loftSS%
		) ELSE (
		START /d "%local%"  %sunSS% -s
		TIMEOUT /t 1800 /NOBREAK 
		TASKKILL /f /im %sunSS%
		)
	)
	
	IF /I "%hr%" EQU "10" ( rem This is setup for 11:30 Service
		START /WAIT /d "%local%" Appload.exe 
		
		IF /I %test% == %loft% (
		START /d "%local%" %loftSS% -s
		TIMEOUT /t 1800 /NOBREAK 
		TASKKILL /f /im %loftSS%
		) ELSE (
		START /d "%local%"  %sunSS% -s
		TIMEOUT /t 1800 /NOBREAK 
		TASKKILL /f /im %sunSS%
		)
	)

	IF /I "%hr%" EQU "12" ( rem This is the end of Sunday Services
		IF /I %test% == %loft% (
		START /d "%local%" %loftSS% -s
		TIMEOUT /t 2700 /NOBREAK 
		TASKKILL /f /im %loftSS%
		) ELSE (
		START /d "%local%"  %sunSS% -s
		TIMEOUT /t 2700 /NOBREAK 
		TASKKILL /f /im %sunSS%
		)
		START nircmd.exe monitor off
	)	
)
rem ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Wednesday~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

IF /I "%dow%" EQU "Wednesday" (
	IF /I "%hr%" EQU "17" ( rem this is Wednesday setup for life groups
		START /WAIT /d "%local%" Appload.exe 
		
		IF /I %test% == %loft% (
		START /d "%local%" %loftSS% -s
		TIMEOUT /t 60
		TASKKILL /f /im %loftSS%
		) ELSE (
		START /d "%local%"  %wedSS% -s
		)
		
	)
	
	IF /I "%hr%" EQU "18" ( rem this is Wednesday setup
		START /WAIT /d "%local%" Appload.exe 
		
		IF /I %test% == %loft% (
		START /d "%local%" %loftSS% -s
		TIMEOUT /t 300 /NOBREAK 
		TASKKILL /f /im %loftSS%
		) ELSE (
		START /d "%local%"  %wedSS% -s
		TIMEOUT /t 300 /NOBREAK 
		TASKKILL /f /im %wedSS%
		)
	)
	
	IF /I "%hr%" EQU "19" ( rem this is the end of Wednesday Service
		IF /I %test% == %loft% (
		START /d "%local%" %loftSS% -s
		TIMEOUT /t 2700 /NOBREAK 
		TASKKILL /f /im %loftSS%
		) ELSE (
		START /d "%local%"  %wedSS% -s
		TIMEOUT /t 2700 /NOBREAK 
		TASKKILL /f /im %wedSS%
		)
		START nircmd.exe monitor off
	)
)	

rem ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Tuesday~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

IF /I "%dow%" EQU "Tuesday" (
	IF /I "%hr%" EQU "17" ( rem this is Tuesday PM
		START /WAIT /d "%local%" Appload.exe 
	)
	
	IF /I "%hr%" EQU "08" ( rem this is Tuesday AM
		START /WAIT /d "%local%" Appload.exe 
	)
)

rem ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Friday~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

IF /I "%dow%" EQU "Friday" (
	IF /I "%hr%" EQU "18" ( rem this is Friday PM
		START /WAIT /d "%local%" Appload.exe 
		START /d "%local%" %hr%SS.exe -s
		TIMEOUT /t 900 /NOBREAK 
		TASKKILL /f /im %hr%SS.exe
	)
)
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#AutoIt3Wrapper_Res_Comment=Fellowship One Check- in Automation
#AutoIt3Wrapper_Res_Description=This program automates Fellowship One Check-in setup with command line input or XML data
#AutoIt3Wrapper_Res_Fileversion=1.0.0.8
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Jonathan Anderson - The Ark Church 2016
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Field=Made By|Jonathan Anderson
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.2
	Author:         Jonathan Anderson
	Script Function: Automating Fellowship One App on Windows machines
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <MsgBoxConstants.au3>
#include <UIAWrappers.au3>
#include <Array.au3>
#include <String.au3>
#include <Date.au3>
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("MustDeclareVars", 1)

; Variable Dec Section
Global $sendcode = False ;var for which window function
Global $code = ""; for SendCode function
Global $mode = ""; for setting mode of application
Global $setmode = False ; var for which window function
Global $nodefault = False ; var for which window function
Global $setupCompleted = False
Global $assisted = False ; var for which window function
Global $CodeCount = 0 ; counts how many times a code was sent to the code screen and will detect a bad code
Global $comp = @ComputerName ; Computer Name For Gettign Station Info
Global $hour = @HOUR ; For code and mode selection in getSettings Function
Global $day = _DateDayOfWeek(@WDAY); For the day
Global $errorCount = 0
Const $F1 = "C:\Fellowship One Check-in 2.6\AppStart.exe"; for CheckWindow function
Const $F1Updater = "AppStart.exe"
Const $F1Win7 = "C:\FT\Fellowship One Check-in 2.5\AppStart.exe"
Const $process = "FellowshipTech.Application.Windows.CheckIn.exe"; for CheckProcess function


#Region ; This region contains all functions~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Func CheckWindow() ; Looks at all the readable text in the current window and decides which window it's in so it knows the action to perform.

	Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")

	For $count = 0 To 5 Step +1

		;MsgBox($MB_SYSTEMMODAL, "", "Loop Count:" & $count)
		If WinExists("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church") Then

			If WinActivate($han, "") Then
				;MsgBox($MB_SYSTEMMODAL, "", "Exists and active")
				ExitLoop
			Else
				WinActivate($han, "")
				ErrorSearch()
				;WinWaitActive($han, "The Ark Church", 5)
				;MsgBox($MB_SYSTEMMODAL, "", "Exists and not active")
			EndIf

		Else
			;MsgBox($MB_SYSTEMMODAL, "", "Wait: " & $count, 2)
			WinWaitActive("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church", 5)
		EndIf

		If $count = 5 Then
			MsgBox($MB_SYSTEMMODAL, "", "Problem with Check-in" & @CRLF & "No Network or Slow Internet!" & @CRLF & "Contact IT department please", 10)
			ExitLoop
		EndIf

	Next

EndFunc   ;==>CheckWindow

Func Checkprocess() ;checks to see if the FellowShip One App is loaded and if it isnt it loads it.

	Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")

	If Not ProcessExists($process) And Not ProcessExists($F1Updater) Then

		If @OSVersion = "WIN_7" Then
			Run($F1Win7)
		Else

			Run($F1) ; checks if process exists.. If not, it will Run the process
			;MsgBox($MB_SYSTEMMODAL, "", "Opening App")
		EndIf
	Else
		ProcessClose($process) ; added because its best to just quit the proces beczuse of error's
		ProcessClose($F1Updater)
		Sleep(1000)
		If @OSVersion = "WIN_7" Then
			Run($F1Win7)
		Else
			Run($F1)
			Sleep(1000)
			WinWaitActive("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church", 5)
			Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")
			;MsgBox($MB_SYSTEMMODAL, "", "Not Opening")
			WinActivate($han, "")
		EndIf
	EndIf

EndFunc   ;==>Checkprocess

Func SendCode($activity) ;for sendig the activity code. Must pass the activity variable. Shouldnt be sent More than once or theres trouble add handling

	Sleep(1000)
	Local $oP0 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	_UIA_Action($oP0, "setfocus")
	_UIA_action(".mainwindow", "sendkeys", $activity)

EndFunc   ;==>SendCode

Func SelfClick() ; For Setting self mode

	;For Clicking the Self Button
	Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	;_UIA_Action($oP1, "setfocus")
	Local $oP0 = _UIA_getObjectByFindAll($oP1, "Title:=Self  Check-in;controltype:=UIA_CheckBoxControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
	;_UIA_Action($oP0, "setfocus")
	_UIA_action($oP0, "leftclick")

	; For Clicking the next button
	Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	;_UIA_Action($oP1, "setfocus")
	Local $oP0 = _UIA_getObjectByFindAll($oP1, "Name:=Next;controltype:=UIA_ButtonControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
	_UIA_action($oP0, "leftclick")

EndFunc   ;==>SelfClick

Func AssistClick() ; For selecting assisted mode and setting rapid

	;This section will click assisted
	Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	;_UIA_Action($oP1, "setfocus")
	Local $oP0 = _UIA_getObjectByFindAll($oP1, "Title:=Assisted Check-in;controltype:=UIA_CheckBoxControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
	;_UIA_Action($oP0, "setfocus")
	_UIA_Action($oP0, "leftclick")

	;This Section will enable rapid
	Local $oP2 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	;_UIA_Action($oP2, "setfocus")
	Local $oP1 = _UIA_getObjectByFindAll($oP2, "Title:=What type of check-in would you like?;controltype:=UIA_PaneControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	;_UIA_Action($oP1, "setfocus")
	Local $oP0 = _UIA_getObjectByFindAll($oP1, "Title:=Enable Rapid;controltype:=UIA_CheckBoxControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
	_UIA_Action($oP0, "leftclick")

	; This section will check the next Button
	Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	;_UIA_Action($oP1, "setfocus")
	Local $oP0 = _UIA_getObjectByFindAll($oP1, "Name:=Next;controltype:=UIA_ButtonControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
	_UIA_action($oP0, "leftclick")

EndFunc   ;==>AssistClick

#cs Added to self click for simpler code. Remove comment if necessary to add function back
	Func NextClick() ; For Clicking Next Button After Self Activity--- should move into the selfclick()

	Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=SwitchForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	_UIA_Action($oP1, "setfocus")
	Local $oP0 = _UIA_getObjectByFindAll($oP1, "Name:=Next;controltype:=UIA_ButtonControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
	_UIA_action($oP0, "leftclick")

	EndFunc   ;==>NextClick
#ce

Func DefaultActivityClick() ; The Default Activity Screen and Options Set

	; This will click the Best fit option
	Local $oP2 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=DefaultActivityForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	Local $oP1 = _UIA_getObjectByFindAll($oP2, "Title:=Select default activity and time for checkin.;controltype:=UIA_PaneControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	;_UIA_Action($oP1, "setfocus")
	Local $oP0 = _UIA_getObjectByFindAll($oP1, "Title:=Best Fit;controltype:=UIA_CheckBoxControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
	_UIA_Action($oP0, "leftclick")

	; This Will Click Start Check-in
	Local $oP1 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=DefaultActivityForm;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
	Local $oP0 = _UIA_getObjectByFindAll($oP1, "Name:=Start Check-in;controltype:=UIA_ButtonControlTypeId;class:=WindowsForms10.BUTTON.app.0.*", $treescope_children)
	_UIA_action($oP0, "leftclick")

EndFunc   ;==>DefaultActivityClick

Func WhichWindow() ; Finding what window the program is in Activity, Set Mode, Code, Etc... ------Needs Error Handling! Either inside this function or from the calling actions-------

	Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")
	Local $text = WinGetText($han, ""); for getting text to find which window we're in
	Local $tArray = _StringExplode($text, @LF); for splitting text into an array
	Local $iIndex = _ArraySearch($tArray, "Activity Code:", Default); for searching the array

	If @error Then; check for self assisted window
		Local $iIndex = _ArraySearch($tArray, "Next", Default)

		If @error Then; check for Default activity No Default button. This screen has the Best Fit Button
			Local $iIndex = _ArraySearch($tArray, "No Default", Default)

			If @error Then
				Local $iIndex = _ArraySearch($tArray, "Scan your card to check-in", Default)

				If @error Then

					Local $iIndex = _ArraySearch($tArray, "Return All Household Members", Default)

					If @error Then

						Local $iIndex = _ArraySearch($tArray, "Application Setup", Default)

						If @error Then

							ErrorSearch()

							MsgBox($MB_SYSTEMMODAL, "", "I Cant find screen Please contact IT department and let them know.", 15)

						Else
							;MsgBox($MB_SYSTEMMODAL, "", "Setup Screen" & $han); for Testing purpose
							;Need to write function to deal with this screen
						EndIf
					Else
						;MsgBox($MB_SYSTEMMODAL,"","Assist" & $han); for testing purpose
						$setupCompleted = True
					EndIf
				Else
					;MsgBox($MB_SYSTEMMODAL,"","Card", $han); for testing purpose
					$setupCompleted = True
				EndIf
			Else
				;MsgBox($MB_SYSTEMMODAL, "", "No Default"); for testing purpose
				$nodefault = True
				$sendcode = False
				$setmode = False
			EndIf
		Else
			;MsgBox($MB_SYSTEMMODAL, "", "Set mode"); for testing purpose
			$setmode = True
			$sendcode = False
			$nodefault = False
		EndIf
	Else
		;MsgBox($MB_SYSTEMMODAL,"","Code"); for testing purpose
		; at main code screen enter code
		$sendcode = True
		$setmode = False
		$nodefault = False
	EndIf

EndFunc   ;==>WhichWindow

Func ErrorSearch() ; This function will find error windows as they come up then perform some action. to be used in later versions

	; this part is setting up the main windoe handle for a given error, subsequent windows might have a differernt class an need different values nested in IF statements
	Local $han = WinGetHandle("Error", "remote") ; This is specific to a window having problems with no internet connection
	Local $text = WinGetText($han, ""); for getting text to find which window we're in
	Local $tArray = _StringExplode($text, @LF); for splitting text into an array
	Local $iIndex = _ArraySearch($tArray, "OK", Default); for searching the array


	If @error Then; Chenge search text value to look for another window.

		Local $han = WinGetHandle("Connection Error", "OK") ; This is specific to a window having problems with no internet connection
		Local $text = WinGetText($han, ""); for getting text to find which window we're in
		Local $tArray = _StringExplode($text, @LF); for splitting text into an array
		Local $iIndex = _ArraySearch($tArray, "OK", Default); for searching the array

		If @error Then

			Local $han = WinGetHandle("ApplicationUpdater", "&Close program") ; This is specific to a window having problems with no internet connection
			Local $text = WinGetText($han, ""); for getting text to find which window we're in
			Local $tArray = _StringExplode($text, @LF); for splitting text into an array
			Local $iIndex = _ArraySearch($tArray, "&Close program", Default); for searching the array

			If @error Then

				;MsgBox($MB_SYSTEMMODAL,"Error", "No Screen Found", 5)

			Else
				MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One Application Updater, Please contact IT" & @CRLF & "Trying Again", 15)
				$errorCount = $errorCount + 1
				Checkprocess()

				If $errorCount > 3 Then
					MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 15)
					Exit
				EndIf
			EndIf
		Else
			MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT" & @CRLF & "Trying Again", 15)
			$errorCount = $errorCount + 1
			Checkprocess()

			If $errorCount > 3 Then
				MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 15)
				Exit
			EndIf
		EndIf
	Else
		; for testing purpose.
		;Local $iIndex = _ArraySearch($tArray, "Activity Code:",Default)
		MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT" & @CRLF & "Trying Again", 15)
		$errorCount = $errorCount + 1
		Checkprocess()

		If $errorCount > 3 Then
			MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 15)
			Exit
		EndIf; at this point an email should be sent or some lan message to a device
		; _ArrayDisplay($tArray, "Works", Default) ; This line will display all the text of a screen for testing
	EndIf

EndFunc   ;==>ErrorSearch

Func Setup(); This is the main setup logic used

	If $code = "" Or $mode = "" Then

		MsgBox($MB_SYSTEMMODAL, "Error", "No settings for Service. Please Contact the IT Department", 15)
		Run(@ScriptDir & "\" & $hour & "SS.exe")
		Exit

	Else


		While $setupCompleted = False ; The Main section for stepping thru the different screens and making decisions based on theh screen

			WinWaitActive("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church", 5)
			;CheckWindow()
			ErrorSearch()
			WhichWindow()
			; Get Computer Setup Information here

			If $sendcode = True Then; decide which code

				$CodeCount = $CodeCount + 1;------------set counter and if statement  for error handleing, if code 3 times wrong code look up code again and try inside this loop, else exit with error code--------------
				If $CodeCount < 4 Then

					SendCode($code)
					$sendcode = False
					Sleep(1000)
					;MsgBox($MB_SYSTEMMODAL, "", $CodeCount)

				Else
					MsgBox($MB_SYSTEMMODAL, "", "Invalid Code. Please Contact Jonathan or Nate on chanel 7 and let them Know", 15)
					$setupCompleted = True
				EndIf


			ElseIf $setmode = True Then; Check for the mode of setup and call the appropiate function

				Select
					Case $mode = "Self"
						SelfClick()

					Case $mode = "Assisted"
						AssistClick()
						DefaultActivityClick()

				EndSelect

				$setmode = False

			ElseIf $nodefault = True Then ;  this will bve nested in the assisted activity actions but left here for checking
				DefaultActivityClick()
				$nodefault = False
				$setupCompleted = False

			ElseIf $setupCompleted = True Then
				ExitLoop
			Else
				MsgBox($MB_SYSTEMMODAL, "", "Problem with Setup, please contact Jonathan or Nate.", 15)
				ExitLoop
			EndIf

		WEnd

	EndIf
EndFunc   ;==>Setup

Func GetSettings($hour, $comp, $day) ; This will get the code and mode to set fo the computer
	Local $cNode = "" ; This will be used for the Current Node Selected
	Local $oXML = ObjCreate("Microsoft.XMLDOM"); Create and XML Object
	$oXML.load(@ScriptDir & "\eSch.xml")

;~~~~~~~~~~~~~~~~~~~~~~ Select the correct service day node~~~~~~~~~~~~~~~~~~~~~~
	Local $nodeDays = $oXML.selectNodes("//Service")

	For $nodeDay In $nodeDays

		Local $nodeProperty = $nodeDay.getAttribute("Day")

		If $nodeProperty = $day Then
			$cNode = $nodeDay
			;MsgBox(0, "Atrib", $nodeProperty); For Testing

;~~~~~~~~~~~~~~~~~~~~Use the service day node to find the hour node~~~~~~~~~~~~~~~~~~~~~~~~
			Local $nodeHours = $cNode.childNodes

			For $nodeHour In $nodeHours

				Local $nodeProperty = $nodeHour.getAttribute("Hour")

				If $nodeProperty = $hour Then
					$cNode = $nodeHour
					;MsgBox(0, "Atrib", $nodeProperty); For Testing

;~~~~~~~~~~~~~~~~~~~~Use hour node to find computer~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

					Local $nodeComputers = $cNode.childNodes

					For $nodeComputer In $nodeComputers

						Local $nodeProperty = $nodeComputer.getAttribute("Station")

						If $nodeProperty = $comp Then
							$cNode = $nodeComputer
							;MsgBox(0, "Comp", $nodeProperty); For Testing

;~~~~~~~~~~~~~~~~~~~use computer to get the child text~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

							Local $nodeSettings = $cNode.childNodes
							For $nodeSetting In $nodeSettings

								If $nodeSetting.baseName = "Code" Then
									$code = $nodeSetting.text
									;MsgBox(0, "Code If", $nodeSetting.baseName & $code); For Testing
								ElseIf $nodeSetting.baseName = "Mode" Then
									$mode = $nodeSetting.text
									;MsgBox(0, "Mode If", $nodeSetting.baseName & $mode); For Testing
								Else
									MsgBox(0, "Problem", "there was a problem getting information for computer" & $cNode.getAttribute("Station"), 15); For Testing
								EndIf
							Next

						EndIf

					Next

				EndIf

			Next

		EndIf

	Next
EndFunc   ;==>GetSettings
#EndRegion ; This region contains all functions~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#Region ;-------- This Section will accept command line parameters if passed and if not the lok thru the XML for settings----------------------------------

If $CmdLine[0] = 0 Then

	GetSettings($hour, $comp, $day)
	;MsgBox(0, "No params", "Nothing exist so check XML")
Else

	If $CmdLine[0] >= 1 Then
		;MsgBox(0, "code", "Exist")

		If $CmdLine[0] >= 2 Then
			;MsgBox(0, "code param", "Exist")
			chkParam1()
		Else
			;MsgBox(0, "code", " No code set, using default")
		EndIf

	EndIf

	If $CmdLine[0] >= 3 Then
		;MsgBox(0, "mode", "Exist")

		If $CmdLine[0] >= 4 Then
			;MsgBox(0, "mode param", "Exist")
			chkParam2()
		Else
			;MsgBox(0, "mode param", "No mode set, using default")
		EndIf

	EndIf
EndIf

Func chkParam1(); this will set command line passed mode and code
	If $CmdLine[1] = "code" Then
		; if the parameter contains anything
		If $CmdLine[2] <> "" Then
			$code = $CmdLine[2]
			;MsgBox(0, "code", "You did it!" & @LF & $code)
		EndIf
	EndIf

	; Checking parameter2
	If $CmdLine[1] = "mode" Then
		; if the parameter contains anything
		If $CmdLine[2] <> "" Then
			$mode = $CmdLine[2]
			;MsgBox(0, "mode", "You did it!" & @LF & $mode)
		EndIf
	EndIf
EndFunc   ;==>chkParam1

Func chkParam2(); this will set command line passed mode and code

	; Checking parameter 1
	If $CmdLine[3] = "code" Then
		; if the parameter contains anything
		If $CmdLine[4] <> "" Then
			$code = $CmdLine[4]
			;MsgBox(0, "code", "You did it!" & @LF & $code)
		EndIf
	EndIf

	; Checking parameter 2
	If $CmdLine[3] = "mode" And $CmdLine[4] <> "" Then
		; if the parameter contains anything
		$mode = $CmdLine[4]
		;MsgBox(0, "mode", "You did it!" & @LF & $mode)
	EndIf
EndFunc   ;==>chkParam2

#EndRegion ;-------- This Section will accept command line parameters if passed and if not the lok thru the XML for settings----------------------------------

Checkprocess(); seeing if app is running if so close and reopen

WinWaitActive("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church", 30); wait for app to become active

ErrorSearch()

CheckWindow()

Setup()
; ------------------------add sleep and continue error search here------------



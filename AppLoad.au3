#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ark.ico
#AutoIt3Wrapper_Res_Comment=Fellowship One Check- in Automation
#AutoIt3Wrapper_Res_Description=This program automates Fellowship One Check-in setup with command line input or XML data
#AutoIt3Wrapper_Res_Fileversion=1.0.0.31
#AutoIt3Wrapper_Res_LegalCopyright=Jonathan Anderson - The Ark Church 2016
#AutoIt3Wrapper_Res_Field=Made By|Jonathan Anderson
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; *** Start added by AutoIt3Wrapper ***
#include <MsgBoxConstants.au3>
; *** End added by AutoIt3Wrapper ***

#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.2
	Author:         Jonathan Anderson
	Script Function: Automating Fellowship One Check-in on Windows machines
	Template AutoIt script.

	There are some small changes to be made for your church. Any where you se The Ark Church your Church name should be, or whats defined on the app screen
	If any information is needed my email is jonathan.andersonATthearkchurchDOTcom

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <MsgBoxConstants.au3>
#include <UIAWrappers.au3>
#include <Array.au3>
#include <String.au3>
#include <Date.au3>
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("MustDeclareVars", 1)
Opt("TrayAutoPause", 0)

;~ Variable Dec Section
Global $sendcode = False ;var for which window function
Global $code ; for SendCode function
Global $mode ; for setting mode of application
Global $setmode = False ; var for which window function
Global $nodefault = False ; var for which window function
Global $setupCompleted = False
Global $assisted = False ; var for which window function
Global $CodeCount = 0 ; counts how many times a code was sent to the code screen
Global $comp = @ComputerName ; Computer Name For Gettign Station Info
Global $hour = @HOUR ; For code and mode selection in getSettings Function
Global $day = _DateDayOfWeek(@WDAY); For the day
Global $errorCount = 0
Global $setupCount = 0
Const $F1 = "C:\Fellowship One Check-in 2.6\AppStart.exe"; for CheckWindow function in windows 8 and 10 machines
Const $F1Updater = "AppStart.exe"
Const $F1Win7 = "C:\FT\Fellowship One Check-in 2.5\AppStart.exe"; directory in windows 7 machines
Const $process = "FellowshipTech.Application.Windows.CheckIn.exe"; for CheckProcess function

#Region ;~ This region Contains Send Email Settings and Functions ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Func SendMail($sub, $msg)

;~ Variable Dec~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Global $SmtpServer = "smtp.gmail.com" ; address for the smtp-server to use - REQUIRED
	Global $FromName = $comp ; name from who the email was sent
	Global $FromAddress = "youremail.com" ; address from where the mail should come
	Global $ToAddress = "someemail.com" ;multiple addresses use semicolin seperator, destination address of the email - REQUIRED
	Global $Subject = $sub ; subject from the email - can be anything you want it to be
	Global $Body = $msg ; the messagebody from the mail - can be left blank but then you get a blank mail
	Global $CcAddress = "" ; address for cc - leave blank if not needed
	Global $BccAddress = "" ; address for bcc - leave blank if not needed
	Global $Importance = "High" ; Send message priority: "High", "Normal", "Low"
	Global $Username = "emailusername.com" ; username for the account used from where the mail gets sent - REQUIRED
	Global $Password = "yes this is plain text" ; password for the account used from where the mail gets sent - REQUIRED
	Global $IPPort = 465 ; port used for sending the mail
	Global $ssl = 1 ; enables/disables secure socket layer sending - put to 1 if using httpS

	Global $rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $CcAddress, $BccAddress, $Importance, $Username, $Password, $IPPort, $ssl)

	If @error Then
		MsgBox(0, "Error sending message", "Error code:" & @error & "  Description:" & $rc)
	EndIf

EndFunc   ;==>SendMail

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance = "Normal", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)

	Local $objEmail = ObjCreate("CDO.Message")
	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress

	Local $i_Error = 0
	Local $i_Error_desciption = ""

	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress

	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

	$objEmail.Subject = $s_Subject

	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
	EndIf

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer

	If Number($IPPort) = 0 Then $IPPort = 25
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

	;Authenticated SMTP
	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf

	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf

	;Update settings
	$objEmail.Configuration.Fields.Update

	; Set Email Importance
	Switch $s_Importance
		Case "High"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "High"
		Case "Normal"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Normal"
		Case "Low"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Low"
	EndSwitch

	$objEmail.Fields.Update

	; Send the Message
	$objEmail.Send

	If @error Then
		SetError(2)
	EndIf

	$objEmail = ""
EndFunc   ;==>_INetSmtpMailCom

#EndRegion ;~ This region Contains Send Email Settings and Functions ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#Region ; This region contains all functions~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Func CheckWindow() ; Looks at all the readable text in the current window and decides which window it's in so it knows the action to perform.

	Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")

	For $count = 0 To 5 Step +1

		;MsgBox($MB_SYSTEMMODAL, "", "Loop Count:" & $count); for testing
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
			MsgBox($MB_SYSTEMMODAL, "", "Problem with Check-in" & @CRLF & "No Network or Slow Internet!" & @CRLF & "Contact IT department please", 5)
			SendMail("Window Error On: " & $comp, "Slow or No connection " & @CRLF & "Time: " & _Now() & @CRLF & "Check Window Count: " & $count & @CRLF & "At Line #: 165")
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

			Run($F1)
		EndIf
	Else
		ProcessClose($process) ; added because its best to just quit the proces because of error's
		ProcessClose($F1Updater)
		Sleep(1000)

		If @OSVersion = "WIN_7" Then ;Run version spcefif file
			Run($F1Win7)
			Sleep(1500)
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

Func SendCode($activity) ;for sendig the activity code. Must pass the activity variable. rewrite in version 1.0.0.30, needs improvement/restructure got tired but works well
	Local $sentCode = 0
	Local $han = WinGetHandle("[REGEXPCLASS:WindowsForms10.Window.8.app.0.*]", "The Ark Church")
	Local $loop = 0
	Local $goodCount = 0
	Local $badCount = 0
	Local $idkCount = 0
	Local $errorCount = 0
	Local $process = True
	Local $errorMsg = ""

	WinActivate($han, "")
	ControlFocus("", "", "[NAME:activityCodeTextBox]")
	ControlSend("", "", "[NAME:activityCodeTextBox]", $activity)
	$sentCode += 1

	Do

		$loop += 1
		ControlFocus("", "", "[NAME:activityCodeTextBox]")
		Local $setTxt = ControlGetText("", "", "[NAME:activityCodeTextBox]")
		WinActivate($han, "")
		Local $cTextError = ControlGetText("", "", "[NAME:errorMessageLabel]")

		If $setTxt = $activity Then ; text matched what i sent

			Do
				$goodCount += 1
				Sleep(1000)
				WhichWindow()
				$cTextError = ControlGetText("", "", "[NAME:errorMessageLabel]")
				$setTxt = ControlGetText("", "", "[NAME:activityCodeTextBox]")

;~ 				If someone touched the screen and messed up the code, warn, -1 from loop count,
				If $setTxt <> $activity And $goodCount > 3 Then
					MsgBox($MB_SYSTEMMODAL, "Warning", "Please don't touch Me!" & @CRLF & "I'm Trying to set up", 2)
					$loop -= 1
					ExitLoop
				EndIf

				If $setmode = True Then
					$process = False
				EndIf

			Until $goodCount = 15 Or $process = False Or $cTextError <> ""

		ElseIf $setTxt <> $activity Then

			$badCount += 1
			Select
				Case $badCount = 1
					ControlFocus("", "", "[NAME:activityCodeTextBox]")
					ControlSetText("", "", "[NAME:activityCodeTextBox]", "")
					ControlSend("", "", "[NAME:activityCodeTextBox]", $activity)
					$sentCode += 1
				Case $badCount = 2
					Local $oP0 = _UIA_getObjectByFindAll($UIA_oDesktop, "Title:=;controltype:=UIA_WindowControlTypeId;class:=WindowsForms10.Window.8.app.0.*", $treescope_children)
					_UIA_Action($oP0, "setfocus")
					ControlSetText("", "", "[NAME:activityCodeTextBox]", "")
					_UIA_action(".mainwindow", "sendkeys", $activity)
					$sentCode += 1
				Case Else
					ControlFocus("", "", "[NAME:activityCodeTextBox]")
					ControlSend("", "", "[NAME:activityCodeTextBox]", $activity)
					$sentCode += 1
			EndSelect
		Else
			$idkCount += 1
		EndIf

;~ 		Look at error message label
		If $cTextError <> "" Then

			$errorMsg = $cTextError
;~ 			reset the error message label to null and hide it
			ControlShow("", "", "[NAME:errorMessageLabel]")
			ControlFocus("", "", "[NAME:errorMessageLabel]")
			ControlSetText("", "", "[NAME:errorMessageLabel]", "")
			ControlHide("", "", "[NAME:errorMessageLabel]")
			WinActivate($han, "")
			ControlFocus("", "", "[NAME:activityCodeTextBox]")

			If $setTxt <> $activity Then ; if its not what it is supposed to be because someone touched it
				ControlSetText("", "", "[NAME:activityCodeTextBox]", "")
			EndIf
			$errorCount = $errorCount + 1
		EndIf

		WhichWindow()

	Until $process = False Or $errorCount = 3 Or $badCount > 3 Or $loop = 4

;~ 	MsgBox($MB_SYSTEMMODAL, "Stats", "Loop count: " & $loop & @CRLF & "Good count: " & $goodCount & @CRLF & "Bad count: " & $badCount & @CRLF & "IDK count: " & $idkCount & @CRLF & "Error Count: " & $errorCount & @CRLF & "Sent Code: " & $sentCode, 15)

	If $process = True And $loop = 4 Then
;~ 		MsgBox($MB_SYSTEMMODAL, "", "Invalid Code. Please Contact Jonathan or Nate on channel 7 and let them Know", 5)
		SendMail("Setup Error On: " & $comp, "Invalid Code" & @CRLF & "Time: " & _Now() & @CRLF & "Loop Count: " & $loop & @CRLF & "Times Code Sent to InputBox: " & $sentCode & @CRLF & "At Line #: 296" & @CRLF & "Code: " & $code & @CRLF & "Mode: " & $mode & @CRLF & "Error Msg: " & $errorMsg)
		Exit (1)
	EndIf

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
	Sleep(500)

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

			If @error Then; check for actual self entry screen

				Local $iIndex = _ArraySearch($tArray, "Scan your card to check-in", Default)

				If @error Then; look for assisted entry screen

					Local $iIndex = _ArraySearch($tArray, "Return All Household Members", Default)

					If @error Then ; look if were in the main app settings screen

						Local $iIndex = _ArraySearch($tArray, "Application Setup", Default)

						If @error Then ; I couldn't find any screen at all so I'll check for errors if founf Ill exit in that function

							ErrorSearch()

							;MsgBox($MB_SYSTEMMODAL, "", "I Cant find screen Please contact IT department and let them know.", 5)

						Else
							;MsgBox($MB_SYSTEMMODAL, "", "Setup Screen" & $han); for Testing purpose
							;Need to write function to deal with this screen
							$setupCompleted = True
						EndIf
					Else
						;MsgBox($MB_SYSTEMMODAL,"","Assist" & $han); for testing purpose
						; at assisted entry screen
						$setupCompleted = True
					EndIf
				Else
					;MsgBox($MB_SYSTEMMODAL,"","Card", $han); for testing purpose
					; at self entry screen
					$setupCompleted = True
				EndIf
			Else
				;MsgBox($MB_SYSTEMMODAL, "", "No Default"); for testing purpose
				; at default activity selection screen
				$nodefault = True
				$sendcode = False
				$setmode = False
			EndIf
		Else
			;MsgBox($MB_SYSTEMMODAL, "", "Set mode"); for testing purpose
			; at set mode screen
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

Func ErrorSearch() ; This function will find error windows as they come up then perform some action

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
				MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One Application Updater, Please contact IT" & @CRLF & "Trying Again", 5)
				$errorCount = $errorCount + 1
				Checkprocess()

				If $errorCount > 3 Then
					SendMail("Error Found On: " & $comp, "Connection Error - Application Updater Error Caught" & @CRLF & "Time: " & _Now() & @CRLF & "Error Count: " & $errorCount & @CRLF & "At Line #: 392")
					MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 5)
					Exit (1)
				EndIf
			EndIf
		Else
			MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT" & @CRLF & "Trying Again", 5)
			$errorCount = $errorCount + 1
			Checkprocess()

			If $errorCount > 3 Then
				SendMail("Error Found On: " & $comp, "Connection Error - No Internet? E-Mail?" & @CRLF & "Time: " & _Now() & @CRLF & "Error Count: " & $errorCount & @CRLF & "At Line #: 403")
				MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 5)
				Exit (1)
			EndIf
		EndIf
	Else
		; for testing purpose.
		;Local $iIndex = _ArraySearch($tArray, "Activity Code:",Default)
		MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT" & @CRLF & "Trying Again", 5)
		$errorCount = $errorCount + 1
		Checkprocess()

		If $errorCount > 3 Then
			SendMail("Error Found On: " & $comp, "Remote Server error message caught" & @CRLF & "Time: " & _Now() & @CRLF & "Error Count: " & $errorCount & @CRLF & "At Line #: 415")
			MsgBox($MB_SYSTEMMODAL, "Error Found", "Error with Fellowship One, Please contact IT", 5)
			Exit (1)
		EndIf; at this point an email should be sent or some lan message to a device
		; _ArrayDisplay($tArray, "Works", Default) ; This line will display all the text of a screen for testing
	EndIf

EndFunc   ;==>ErrorSearch

Func Setup(); This is the main setup logic used

	If $code = "" Or $mode = "" Then

		SendMail("Setup Error On: " & $comp, "No Code or Mode set" & @CR & "Time: " & _Now() & @CR & "Error Count: " & $errorCount & @CR & "At Line #: 428")
		MsgBox($MB_SYSTEMMODAL, "Error", "No settings for Service. Contact IT or enter the correct code.", 5)
;~ 		Run(@ScriptDir & "\" & "Help.exe") ;~~~~~~~~~~~~~~ ADD IT HELP SCREEN HERE~~~~~~~~~~~~~~~~~~~
		Exit (1)

	ElseIf $code = "0" Then

		Exit

	Else

		While $setupCompleted = False ; The Main section for stepping thru the different screens and making decisions based on theh screen

			ErrorSearch(); look for error after every entry
			WhichWindow(); look for next window

			If $sendcode = True Then; decide which code

				SendCode($code)
				$sendcode = False

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
				;MsgBox($MB_SYSTEMMODAL, "", "Problem with Setup, please contact Jonathan or Nate.", 5)
				Sleep(500)
				$setupCount = $setupCount + 1
				If $setupCount > 5 Then
					SendMail("Setup Error On: " & $comp, "Problem with setup" & @CRLF & "Time: " & _Now() & @CRLF & "Setup Count: " & $setupCount & @CRLF & "At Line #: 485" & @CRLF & "Code: " & $code & @CRLF & "Mode: " & $mode)
					Exit (1)
				EndIf
			EndIf

		WEnd

	EndIf
EndFunc   ;==>Setup

Func GetSettings($hour, $comp, $day) ; This will get the code and mode to set fo the computer
	Local $cNode = "" ; This will be used for the Current Node Selected
	Local $oXML = ObjCreate("Microsoft.XMLDOM"); Create and XML Object
	$oXML.load(@ScriptDir & "\eSch.xml")

;~~~~~~~~~~~~~~~~~~~~~~SPECIAL SERVICE SECTION~~~~~~~~~~~~~~~~Look For a Special service and if yes then get settings~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Local $nodeSpecial = $oXML.selectNodes("//Service")

	For $special In $nodeSpecial

		Local $nodeProperty = $special.getAttribute("Special")

		If $nodeProperty = "Yes" Then

			;MsgBox(0,"Special Found", "Node: " & $special)
			$cNode = $special

;~~~~~~~~~~~~~~~~~~~~Use special node to find computer~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			Local $nodeComputers = $cNode.childNodes
			;$count = 0
			For $nodeComputer In $nodeComputers
				;$count = $count + 1
				Local $nodeProperty = $nodeComputer.getAttribute("Station")

				;MsgBox(0, "For Comp Loop", $nodeProperty & "Node comp count is : " & $nodeComputer & @CRLF & "For Count is: " & $count); For Testing

				If @error Then
					;MsgBox(0,"Error", @exitCode)

				ElseIf $nodeProperty = $comp Then
					$cNode = $nodeComputer
					;MsgBox(0, "Comp else", $nodeProperty); For Testing

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
							;MsgBox(0, "Problem", "there was a problem getting information for computer" & $cNode.getAttribute("Station"), 15); For Testing
						EndIf
					Next

				EndIf

			Next

		Else
			;MsgBox(0,"No Special", "Node: " & $special)
		EndIf
	Next
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End of special services sections~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
									;MsgBox(0, "Problem", "there was a problem getting information for computer" & $cNode.getAttribute("Station"), 15); For Testing
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

	;Deal with in between service for manually running program without passing command line code
	If $hour = "08" And $day = "Sunday" Then $hour = "07"
	If $hour = "10" And @MIN < 30 Then $hour = "09"
	If $hour = "11" Then $hour = "10"
	if $hour = "19" And $day = "Wednesday" Then $hour = "18"
	if $hour = "18" And $day = "Tuesday" Then $hour = "17"

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

Setup()

Exit (0)



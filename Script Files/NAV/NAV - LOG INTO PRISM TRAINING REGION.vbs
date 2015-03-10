'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - LOG INTO PRISM TRAINING REGION.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

'VARIABLES AND CALCULATIONS----------------------------------------------------------------------------------------------------
'PRISM training uses the current month as part of the password. This figures out what it needs to be.
date_for_PW = datepart("m", date) 
If len(date_for_PW) = 1 then date_for_PW = "0" & date_for_PW

'Connects to BlueZone
EMConnect ""

EMReadScreen ADMNET_check, 6, 1, 2
If ADMNET_check <> "ADMNET" then script_end_procedure("You are not in ADMNET (main STATE OF MN screen). The script will now stop.")

EMWriteScreen "cicsdt4", 12, 61
transmit
EMWaitReady 0, 0 'waits as the script might hang
EMWriteScreen "pwcst05", 12, 21
EMWriteScreen "Train#" & date_for_PW, 13, 21
transmit
EMSendKey "QQT4"
transmit

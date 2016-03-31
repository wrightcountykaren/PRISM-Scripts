'Gathering stats
name_of_script = "Action - CP NAME CHANGE.vbs"
start_time = timeer
STATS_Counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'End of stats block 

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
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

'the script---------------------------------------------------------------------------------------------------------------
BeginDialog Name_change_dialog, 0, 0, 191, 110, "CP Name Change"
  EditBox 80, 5, 100, 15, Prism_case_number
  EditBox 80, 25, 100, 15, New_name
  EditBox 80, 45, 100, 15, reason_change
  EditBox 80, 65, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 90, 50, 15
    CancelButton 140, 90, 50, 15
  Text 5, 50, 65, 15, "Reason for change:"
  Text 5, 30, 70, 15, "CP New Last Name:"
  Text 5, 70, 65, 15, "Worker Signature:"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog

' connects to Bluezone
EMConnect ""

'Navigates to CAST
EMwritescreen "cast", 21, 08

'Grabs the case number
call PRISM_case_number_finder(Prism_case_number)

'Calls the dialog
Dialog Name_change_dialog

'if cancel button is pressed script is canceled
If buttonpressed = 0 then stopscript

'Navigates to CPDE
call navigate_to_prism_screen ("CPDE")

'Hits transmit
EMsendkey "<enter>"

'Enters "M" to modify
EMwritescreen "M", 3, 29

'Clears last name
EMWritescreen "__________________", 8, 8

'Hits transmit
transmit

'Hits tranmit
transmit

'Enters "M" to modify
EMwritescreen "M", 3, 29

'Enters new last name from dialog
EMwritescreen new_name, 8,8

'hits transmit
transmit

'hits transmit
transmit

'Navigates to CAAD
call navigate_to_prism_screen("CAAD")

'Enters "M" to modify
EMwritescreen "M", 8,5

'hits transmit
transmit

emsetcursor 16,4

'Enters info for CAAD note
call write_bullet_and_variable_in_caad("Updated CP Name to", New_name)


'enters info on CAAD note
call write_bullet_and_variable_in_caad("Reason for change", reason_change)

'enters CSO signature
call write_variable_in_caad(worker_signature)

'hits transmit
transmit


call script_end_procedure ("")












 

 
















 

 




'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MAIN MENU.vbs"
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

'DIALOGS---------------------------------------------------------------------------
BeginDialog NOTES_main_menu_dialog, 0, 0, 306, 100, "NOTES main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 250, 80, 50, 15
    PushButton 5, 5, 50, 10, "Client contact", NOTES_client_contact_button
    PushButton 5, 20, 90, 10, "Date of the hearing (expro)", NOTES_date_of_hearing_expro_button
    PushButton 5, 35, 95, 10, "Date of the hearing (judicial)", NOTES_date_of_hearing_judicial_button
    PushButton 5, 50, 50, 10, "Pay or report", NOTES_pay_or_report_button
    PushButton 5, 65, 65, 10, "Quarterly reviews", NOTES_quarterly_reviews_button
  Text 60, 5, 240, 10, "-- Creates a uniform CAAD note for when you have contact with a client."
  Text 100, 20, 200, 10, "-- Date of the hearing template for expro."
  Text 105, 35, 195, 10, "-- Date of the hearing template for judicial."
  Text 60, 50, 240, 10, "-- CAAD note for case noting ''pay or report'' instances."
  Text 75, 65, 225, 10, "-- CAAD note for quarterly review processes."
EndDialog


'THE SCRIPT-----------------------------------------------------------------------------------------------

'Shows the dialog
Dialog NOTES_main_menu_dialog
If buttonpressed = cancel then StopScript
IF ButtonPressed = NOTES_Client_contact_button then call run_from_GitHub(script_repository & "NOTES/NOTES - CLIENT CONTACT.vbs")
IF ButtonPressed = NOTES_date_of_hearing_expro_button then call run_from_GitHub(script_repository & "NOTES/NOTES - DATE OF THE HEARING (EXPRO).vbs")
IF ButtonPressed = NOTES_date_of_hearing_judicial_button then call run_from_GitHub(script_repository & "NOTES/NOTES - DATE OF THE HEARING (JUDICIAL).vbs")
IF ButtonPressed = NOTES_pay_or_report_button then call run_from_GitHub(script_repository & "NOTES/NOTES - PAY OR REPORT.vbs")
IF ButtonPressed = NOTES_quarterly_reviews_button then call run_from_GitHub(script_repository & "NOTES/NOTES - QUARTERLY REVIEWS.vbs")
'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


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

'-----The dialog-----
BeginDialog NOTES_main_menu_dialog, 0, 0, 431, 260, "NOTES main menu dialog"
  ButtonGroup ButtonPressed
    PushButton 10, 10, 50, 10, "Client contact", NOTES_client_contact_button
    PushButton 10, 25, 80, 10, "Court Order Requested", NOTES_court_order_requested_button
    PushButton 10, 50, 50, 10, "CSENET Info", NOTES_CSENET_button
    PushButton 10, 65, 90, 10, "Date of the hearing (expro)", NOTES_date_of_hearing_expro_button
    PushButton 10, 80, 95, 10, "Date of the hearing (judicial)", NOTES_date_of_hearing_judicial_button
    PushButton 10, 95, 70, 10, "No Pay Months 1-4", NOTES_no_pay_months_button
    PushButton 10, 110, 50, 10, "Pay or report", NOTES_pay_or_report_button
    PushButton 10, 125, 65, 10, "Quarterly reviews", NOTES_quarterly_reviews_button
    PushButton 10, 140, 50, 10, "ROP Invoice", NOTES_ROP_invoice_button
    PushButton 10, 165, 50, 10, "SOP Invoice", NOTES_SOP_invoice_button
    PushButton 10, 190, 105, 10, "Waiver of Personal Service", NOTES_waiver_of_personal_service_button
    CancelButton 375, 240, 50, 15
  Text 65, 10, 240, 10, "-- Creates a uniform CAAD note for when you have contact with a client."
  Text 95, 25, 325, 20, "-- NEW 09/2015!! -- Creates B0170 CAAD note for requesting a court order, which also creates a work list to remind the worker to check the status of the court order request."
  Text 65, 50, 350, 10, "-- NEW 09/2015!! -- Creates T0111 CAAD note script with text copied from the INTD screen."
  Text 105, 65, 200, 10, "-- Date of the hearing template for expro."
  Text 110, 80, 195, 10, "-- Date of the hearing template for judicial."
  Text 85, 95, 340, 10, "-- NEW 09/2015!! -- Creates CAAD note for documenting non-payment enforcement actions."
  Text 65, 110, 240, 10, "-- CAAD note for case noting ''pay or report'' instances."
  Text 80, 125, 225, 10, "-- CAAD note for quarterly review processes."
  Text 70, 140, 350, 20, "-- NEW 09/2015!! -- Creates CAAD note that the Service of Process invoice was received, details about the service, and if the invoice is OK to pay."
  Text 70, 165, 350, 20, "-- NEW 09/2015!! -- Creates CAAD note that the Service of Process invoice was received, details about the service, and if the invoice is OK to pay."
  Text 120, 190, 285, 15, "-- NEW 09/2015!! -- Creates CAAD note of the date a CP signed the waiver of personal service document."
EndDialog


'THE SCRIPT-----------------------------------------------------------------------------------------------

'Shows the dialog
Dialog NOTES_main_menu_dialog
If buttonpressed = cancel then StopScript
IF ButtonPressed = NOTES_Client_contact_button then call run_from_GitHub(script_repository & "NOTES/NOTES - CLIENT CONTACT.vbs")
IF ButtonPressed = NOTES_court_order_requested_button THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - COURT ORDER REQUEST.vbs")
IF ButtonPressed = NOTES_CSENET_button THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CSENET INFO.vbs")
IF ButtonPressed = NOTES_date_of_hearing_expro_button then call run_from_GitHub(script_repository & "NOTES/NOTES - DATE OF THE HEARING (EXPRO).vbs")
IF ButtonPressed = NOTES_date_of_hearing_judicial_button then call run_from_GitHub(script_repository & "NOTES/NOTES - DATE OF THE HEARING (JUDICIAL).vbs")
IF ButtonPressed = NOTES_no_pay_months_button THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - NO PAYMENT MONTHS ONE-FOUR.vbs")
IF ButtonPressed = NOTES_pay_or_report_button then call run_from_GitHub(script_repository & "NOTES/NOTES - PAY OR REPORT.vbs")
IF ButtonPressed = NOTES_quarterly_reviews_button then call run_from_GitHub(script_repository & "NOTES/NOTES - QUARTERLY REVIEWS.vbs")
IF ButtonPressed = NOTES_ROP_invoice_button THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - ROP DETAIL.vbs")
IF ButtonPressed = NOTES_SOP_invoice_button THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - SOP INVOICE.vbs")
IF ButtonPressed = NOTES_waiver_of_personal_service_button THEN CALL run_from_GitHub (script_repository & "NOTES/NOTES - WAIVER OF PERSONAL SERVICE.vbs")

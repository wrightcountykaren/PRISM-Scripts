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
BeginDialog NOTES_main_menu_dialog, 0, 0, 437, 346, "NOTES main menu dialog"
  ButtonGroup ButtonPressed
    PushButton 10, 50, 50, 10, "Client contact", NOTES_client_contact_button
    PushButton 10, 70, 80, 10, "Court Order Requested", NOTES_court_order_requested_button
    PushButton 10, 90, 50, 10, "CSENET Info", NOTES_CSENET_button
    PushButton 10, 110, 90, 10, "Date of the hearing (expro)", NOTES_date_of_hearing_expro_button
    PushButton 10, 130, 100, 10, "Date of the hearing (judicial)", NOTES_date_of_hearing_judicial_button
    PushButton 10, 190, 70, 10, "No Pay Months 1-4", NOTES_no_pay_months_button
    PushButton 10, 210, 50, 10, "Pay or report", NOTES_pay_or_report_button
    PushButton 10, 230, 70, 10, "Quarterly reviews", NOTES_quarterly_reviews_button
    PushButton 10, 250, 50, 10, "ROP Detail", NOTES_ROP_invoice_button
    PushButton 10, 270, 50, 10, "SOP Invoice", NOTES_SOP_invoice_button
    PushButton 10, 290, 110, 10, "Waiver of Personal Service", NOTES_waiver_of_personal_service_button
    PushButton 10, 10, 40, 10, "Adjustment", NOTES_adjustment_button
    PushButton 10, 30, 80, 10, "Arrears Management", NOTES_Arrears_mgmt_button
    PushButton 10, 170, 100, 10, "MES Financial Docs Sent", NOTES_MES_Fin_docs_button
    PushButton 10, 150, 70, 10, "Intake Docs Rec'd", NOTES_Intake_docs_button
    PushButton 350, 0, 80, 10, "PRISM Scripts in SIR", SIR_button
    CancelButton 380, 320, 50, 20
  Text 90, 30, 300, 10, "-- NEW 2/2016 Creates a CAAD note for documenting an arrears management review."
  Text 90, 70, 330, 20, "-- Creates B0170 CAAD note for requesting a court order, which also creates a work list to remind the worker to check the status of the court order request."
  Text 60, 90, 350, 10, "-- Creates T0111 CAAD note script with text copied from the INTD screen."
  Text 100, 110, 200, 10, "-- Date of the hearing template for expro."
  Text 110, 130, 200, 10, "-- Date of the hearing template for judicial."
  Text 80, 190, 340, 10, "-- Creates CAAD note for documenting non-payment enforcement actions."
  Text 60, 210, 240, 10, "-- CAAD note for case noting ''pay or report'' instances."
  Text 80, 230, 230, 10, "-- CAAD note for quarterly review processes."
  Text 60, 250, 350, 10, "-- Creates CAAD note noting the dates parties signed recognition of parentage."
  Text 60, 270, 350, 20, "-- Creates CAAD note that the Service of Process invoice was received, details about the service, and if the invoice is OK to pay."
  Text 120, 290, 290, 20, "-- Creates CAAD note of the date a CP signed the waiver of personal service document."
  Text 60, 50, 240, 10, "-- Creates a uniform CAAD note for when you have contact with a client."
  Text 50, 10, 300, 10, "-- NEW 1/2016 Creates a CAAD note for documenting adjustments made to the case."
  Text 110, 170, 290, 10, "-- NEW 2/2016 Creates a CAAD note for recording documents sent to the parties."
  Text 80, 150, 280, 10, "-- NEW 2/2016 Creates a CAAD note for recording receipt of intake docs."
EndDialog




'THE SCRIPT-----------------------------------------------------------------------------------------------

DO
	'Shows the dialog
	Dialog NOTES_main_menu_dialog
	If buttonpressed = cancel then StopScript
	IF ButtonPressed = SIR_button THEN CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/PRISMscripts/PRISM%20script%20wiki/Forms/AllPages.aspx")
LOOP UNTIL ButtonPressed <> SIR_button

IF ButtonPressed = NOTES_adjustment_button then call run_from_GitHub(script_repository & "NOTES/NOTES - ADJUSTMENTS.vbs")
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
IF ButtonPressed = NOTES_Arrears_mgmt_button THEN CALL run_from_GitHub (script_repository & "NOTES/NOTES - ARREARS MGMT REVIEW.vbs")
IF ButtonPressed = NOTES_MES_Fin_docs_button THEN CALL run_from_GitHub (script_repository & "NOTES/NOTES - MES FINANCIAL DOCS SENT.vbs")
IF ButtonPressed = NOTES_Intake_docs_button THEN CALL run_from_GitHub (script_repository & "NOTES/NOTES - INTAKE DOCS RECEIVED.vbs")

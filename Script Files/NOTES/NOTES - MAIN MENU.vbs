'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MAIN MENU.vbs"
start_time = timer

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
			"Before contacting Robert Kalb, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Robert Kalb and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Robert will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

'-----The dialog-----
BeginDialog NOTES_main_menu_dialog, 0, 0, 436, 340, "NOTES main menu dialog"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 40, 10, "Adjustment", NOTES_adjustment_button
    PushButton 10, 35, 80, 10, "Arrears Management", NOTES_Arrears_mgmt_button
    PushButton 10, 50, 50, 10, "Client contact", NOTES_client_contact_button
    PushButton 10, 65, 80, 10, "Court Order Requested", NOTES_court_order_request_button
    PushButton 10, 85, 50, 10, "CSENET Info", NOTES_CSENET_button
    PushButton 10, 100, 90, 10, "Date of the hearing (expro)", NOTES_date_of_hearing_expro_button
    PushButton 10, 115, 100, 10, "Date of the hearing (judicial)", NOTES_date_of_hearing_judicial_button
    PushButton 10, 130, 35, 10, "E-Filing", NOTES_efiling_button
    PushButton 10, 145, 50, 10, "Fraud Referral", NOTES_fraud_referral_button
    PushButton 10, 160, 70, 10, "Intake Docs Rec'd", NOTES_Intake_docs_button
    PushButton 10, 175, 80, 10, "IW CAAD and CAWT", NOTES_IW_caad_button
    PushButton 10, 190, 100, 10, "MES Financial Docs Sent", NOTES_MES_Fin_docs_button
    PushButton 10, 205, 70, 10, "No Pay Months 1-4", NOTES_no_pay_months_button
    PushButton 10, 220, 50, 10, "Pay or report", NOTES_pay_or_report_button
    PushButton 10, 235, 70, 10, "Quarterly reviews", NOTES_quarterly_reviews_button
    PushButton 10, 250, 50, 10, "ROP Detail", NOTES_ROP_invoice_button
    PushButton 10, 265, 50, 10, "SOP Invoice", NOTES_SOP_invoice_button
    PushButton 10, 285, 95, 10, "Waiver of Personal Service", NOTES_waiver_of_personal_service_button
    PushButton 350, 5, 80, 10, "PRISM Scripts in SIR", SIR_button
    CancelButton 380, 305, 50, 15
  Text 95, 35, 300, 10, "-- Creates a CAAD note for documenting an arrears management review."
  Text 95, 65, 330, 20, "-- Creates B0170 CAAD note for requesting a court order, which also creates a work list to remind the worker to check the status of the court order request."
  Text 70, 85, 350, 10, "-- Creates T0111 CAAD note script with text copied from the INTD screen."
  Text 100, 100, 200, 10, "-- Date of the hearing template for expro."
  Text 110, 115, 200, 10, "-- Date of the hearing template for judicial."
  Text 85, 205, 340, 10, "-- Creates CAAD note for documenting non-payment enforcement actions."
  Text 65, 220, 240, 10, "-- CAAD note for case noting ''pay or report'' instances."
  Text 85, 235, 230, 10, "-- CAAD note for quarterly review processes."
  Text 65, 250, 350, 10, "-- Creates CAAD note noting the dates parties signed recognition of parentage."
  Text 65, 265, 350, 20, "-- Creates CAAD note that the Service of Process invoice was received, details about the service, and if the invoice is OK to pay."
  Text 110, 285, 305, 10, "-- Creates CAAD note of the date a CP signed the waiver of personal service document."
  Text 65, 50, 240, 10, "-- Creates a uniform CAAD note for when you have contact with a client."
  Text 50, 20, 300, 10, "-- Creates a CAAD note for documenting adjustments made to the case."
  Text 110, 190, 290, 10, "-- NEW 02/2016!! Creates a CAAD note for recording documents sent to the parties."
  Text 85, 160, 280, 10, "-- NEW 02/2016!! Creates a CAAD note for recording receipt of intake docs."
  Text 50, 130, 350, 10, "-- NEW 04/2016!! Template for adding a CAAD note about e-filing."
  Text 95, 175, 280, 10, "-- NEW 04/2016!! Creates CAAD and CAWT about IW."
  Text 65, 145, 350, 10, "-- NEW 04/2016!! Template for adding a CAAD note about a fraud referral."
EndDialog


'THE SCRIPT-----------------------------------------------------------------------------------------------

DO
	'Shows the dialog
	Dialog NOTES_main_menu_dialog
	If buttonpressed = cancel then StopScript
	IF ButtonPressed = SIR_button THEN CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/PRISMscripts/PRISM%20script%20wiki/Forms/AllPages.aspx")
LOOP UNTIL ButtonPressed <> SIR_button

IF ButtonPressed = NOTES_adjustment_button 			then call run_from_GitHub(script_repository & "NOTES/NOTES - ADJUSTMENTS.vbs")
IF ButtonPressed = NOTES_Arrears_mgmt_button			THEN CALL run_from_GitHub (script_repository & "NOTES/NOTES - ARREARS MGMT REVIEW.vbs")
IF ButtonPressed = NOTES_Client_contact_button 			then call run_from_GitHub(script_repository & "NOTES/NOTES - CLIENT CONTACT.vbs")
IF ButtonPressed = NOTES_court_order_request_button		THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - COURT ORDER REQUEST.vbs")
IF ButtonPressed = NOTES_CSENET_button 				THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CSENET INFO.vbs")
IF ButtonPressed = NOTES_date_of_hearing_expro_button 		then call run_from_GitHub(script_repository & "NOTES/NOTES - DATE OF THE HEARING (EXPRO).vbs")
IF ButtonPressed = NOTES_date_of_hearing_judicial_button 	then call run_from_GitHub(script_repository & "NOTES/NOTES - DATE OF THE HEARING (JUDICIAL).vbs")
IF ButtonPressed = NOTES_efiling_button 			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - E-FILING.vbs")
IF ButtonPressed = NOTES_fraud_referral_button			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - FRAUD REFERRAL.vbs")
IF ButtonPressed = NOTES_Intake_docs_button 			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - INTAKE DOCS RECEIVED.vbs")
IF ButtonPressed = NOTES_IW_caad_button				THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - IW CAAD CAWT.vbs")
IF ButtonPressed = NOTES_no_pay_months_button 			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - NO PAYMENT MONTHS ONE-FOUR.vbs")
IF ButtonPressed = NOTES_pay_or_report_button 			then call run_from_GitHub(script_repository & "NOTES/NOTES - PAY OR REPORT.vbs")
IF ButtonPressed = NOTES_quarterly_reviews_button 		then call run_from_GitHub(script_repository & "NOTES/NOTES - QUARTERLY REVIEWS.vbs")
IF ButtonPressed = NOTES_ROP_invoice_button 			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - ROP DETAIL.vbs")
IF ButtonPressed = NOTES_SOP_invoice_button 			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - SOP INVOICE.vbs")
IF ButtonPressed = NOTES_waiver_of_personal_service_button 	THEN CALL run_from_GitHub (script_repository & "NOTES/NOTES - WAIVER OF PERSONAL SERVICE.vbs")
IF ButtonPressed = NOTES_MES_Fin_docs_button 			THEN CALL run_from_GitHub (script_repository & "NOTES/NOTES - MES FINANCIAL DOCS SENT.vbs")



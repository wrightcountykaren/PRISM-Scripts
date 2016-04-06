'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
'Dim url, REQ, FSO					'Declares variables to be good to option explicit users
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

'DIALOGS---------------------------------------------------------------------------
BeginDialog ACTIONS_main_menu_dialog, 0, 0, 381, 300, "ACTIONS Main Menu"
  ButtonGroup ButtonPressed
    PushButton 5, 25, 85, 10, "Affidavit of Service Docs", ACTIONS_affidavit_of_service_button
    PushButton 5, 40, 30, 10, "COLA", ACTIONS_cola_button
    PushButton 5, 55, 60, 10, "DDPL Calculator", ACTIONS_DDPL_CALC_button
    PushButton 5, 70, 80, 10, "Estab NPA DORD Docs", ACTIONS_EST_DORD_NPA_button
    PushButton 5, 85, 80, 10, "Estab PA DORD Docs", ACTIONS_EST_DORD_PA_button
    PushButton 5, 100, 70, 10, "Find Name on CALI", ACTIONS_find_name_on_cali_button
    PushButton 5, 115, 30, 10, "Intake", ACTIONS_intake_button
    PushButton 5, 135, 30, 10, "IW Calc", ACTIONS_iw_calc_button
    PushButton 5, 155, 45, 10, "NCP Locate", ACTIONS_NCP_locate_button
    PushButton 5, 175, 60, 10, "PALC calculator", ACTIONS_PALC_calculator_button
    PushButton 5, 195, 60, 10, "Prorate Support", ACTIONS_prorate_support_button
    PushButton 5, 215, 65, 10, "Redirection Docs", ACTIONS_redirection_docs_button
    PushButton 5, 235, 75, 10, "Unreimb/Unins Docs", ACTIONS_un_un_button
    PushButton 5, 255, 85, 10, "Update Worker Signature", ACTIONS_update_worker_signature_button
    CancelButton 325, 280, 50, 15
    PushButton 300, 5, 75, 10, "PRISM Scripts in SIR", SIR_button
  Text 100, 25, 270, 10, "-- Sends Affidavits of Serivce to multiple participants on the case."
  Text 70, 55, 300, 10, "-- Calculates payments received during a specific date range."
  Text 90, 70, 280, 10, "-- NEW 01/2016!! Generates DORD docs for NPA case."
  Text 90, 85, 280, 10, "-- NEW 01/2016!! Generates DORD docs for PA case."
  Text 80, 100, 215, 10, "-- Searches CALI for a specific CP or NCP."
  Text 40, 115, 330, 15, "-- Creates various documents related to Child Support intake, as well as DORD documents, and enters a note on CAAD."
  Text 70, 175, 230, 10, "-- Calculates voluntary and involuntary amounts from the PALC screen."
  Text 70, 195, 225, 10, "- Calculator for deteremining pro-rated support for partial months."
  Text 75, 215, 290, 10, "-- Creates redirection docs and redirection worklist items."
  Text 85, 235, 290, 10, "-- Prints DORD docs for collecting unreimbursed and unisured expenses."
  Text 95, 255, 280, 10, "-- NEW 04/2016!! -- Allows you to maintain a default signature that loads in all scripts."
  Text 40, 40, 330, 10, "-- NEW 04/2016!!! Leads you through performing a COLA. Adds CAAD note when completed."
  Text 55, 155, 320, 10, "-- NEW 04/2016!!! Walks you through processing an NCP locate."
  Text 40, 135, 325, 10, "-- NEW 04/2016!! Calculator for determining the amount of IW over a given period."
EndDialog

'THE SCRIPT-----------------------------------------------------------------------------------------------

'Shows the dialog
DO
	Dialog ACTIONS_main_menu_dialog
	If buttonpressed = cancel then stopscript
	IF ButtonPressed = SIR_button THEN CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/PRISMscripts/PRISM%20script%20wiki/Forms/AllPages.aspx")
LOOP UNTIL ButtonPressed <> SIR_button

IF ButtonPressed = ACTIONS_affidavit_of_service_button 		THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - AFFIDAVIT OF SERVICE BY MAIL DOCS.vbs")
IF ButtonPressed = ACTIONS_cola_button 				THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - COLA.vbs")
IF ButtonPressed = ACTIONS_DDPL_CALC_button 			THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - DDPL CALCULATOR.vbs")
IF ButtonPressed = ACTIONS_EST_DORD_NPA_button 			THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - EST DORD NPA DOCS.vbs")
IF ButtonPressed = ACTIONS_EST_DORD_PA_button 			THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - ESTB DORD PA DOCS.vbs")
IF ButtonPressed = ACTIONS_find_name_on_cali_button 		THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - FIND NAME ON CALI.vbs")
IF ButtonPressed = ACTIONS_prorate_support_button 		THEN call run_from_GitHub(script_repository & "ACTIONS/ACTIONS - PRORATE SUPPORT.vbs")
IF ButtonPressed = ACTIONS_intake_button 			THEN call run_from_GitHub(script_repository & "ACTIONS/ACTIONS - GENERIC ENFORCEMENT INTAKE.vbs")
IF ButtonPressed = ACTIONS_iw_calc_button			THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - IW CALCULATOR.vbs
IF ButtonPressed = ACTIONS_ncp_locate_button			THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - NCP LOCATE.vbs")
IF ButtonPressed = ACTIONS_PALC_calculator_button 		THEN call run_from_GitHub(script_repository & "ACTIONS/ACTIONS - PALC CALCULATOR.vbs")
IF ButtonPressed = ACTIONS_redirection_docs_button 		THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - DOCS FOR REDIRECT.vbs")
IF ButtonPressed = ACTIONS_un_un_button 			THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - UNREIMBURSED UNINSURED DOCS.vbs")
IF ButtonPressed = ACTIONS_update_worker_signature_button 	THEN CALL run_from_Github(script_repository & "ACTIONS/ACTIONS - UPDATE WORKER SIGNATURE.vbs")

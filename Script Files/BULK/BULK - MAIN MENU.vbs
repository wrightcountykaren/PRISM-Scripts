''GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MAIN MENU.vbs"
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

'DIALOGS---------------------------------------------------------------------------
BeginDialog BULK_main_menu_dialog, 0, 0, 446, 356, "BULK Main Menu"
  ButtonGroup ButtonPressed
    PushButton 0, 30, 60, 10, "CALI to Excel", BULK_cali_to_excel_button
    PushButton 0, 50, 60, 10, "Case Transfer", BULK_case_transfer_button
    PushButton 0, 80, 100, 10, "Companion Case Finder - CP", BULK_cp_companion_case_finder_button
    PushButton 0, 100, 110, 10, "Companion Case Finder - NCP", BULK_ncp_companion_case_finder_button
    PushButton 0, 120, 60, 10, "Evaluate NOCS", BULK_evaluate_nocs_button
    PushButton 0, 140, 90, 10, "Failure POF -- SSA, DFAS", BULK_failure_pof_rsdi_dfas_button
    PushButton 0, 160, 130, 10, "FI Match Not Eligible For Levy Scrubber", BULK_L5000_button
    PushButton 0, 200, 130, 10, "Review Continued Interest Suspension", BULK_M6529_button
    PushButton 0, 180, 110, 10, "Review Case Referred Scrubber", BULK_M8001_button
    PushButton 0, 230, 100, 10, "Review Quarterly Wage Info", BULK_REVIEW_QW_button
    PushButton 0, 270, 140, 10, "Review Pay Plan Recent Payment Activity", BULK_E4111_activity_button
    PushButton 0, 250, 120, 10, "Review Pay Plan - DL is Suspended", BULK_E4111_suspended_button
    CancelButton 390, 330, 50, 20
    PushButton 360, 0, 80, 10, "PRISM Scripts in SIR", SIR_button
  Text 60, 30, 370, 20, "-- This script builds a list in Microsoft Excel of case numbers, function types, program codes, interstate codes, and participant names based on a CALI caseload."
  Text 60, 50, 380, 20, "-- This script allows users to transfer up to 15 cases to as many workers as they need OR to transfer an entire caseload to as many workers as needed."
  Text 100, 80, 260, 10, "--- This script builds a list of companion cases for your CPs on a given CALI."
  Text 110, 100, 260, 10, "-- This script builds a list of companion cases for your NCPs on a given CALI."
  Text 60, 120, 370, 20, "-- This script evaluates D0800 (Review for Notice of Continued Services) worklist items and allows user to send docs."
  Text 90, 140, 350, 10, "-- Clears E0014 (Failure Notice to POF review) worklist item when income is from RSDI (US Treasury) or DFAS."
  Text 130, 160, 290, 20, "-- NEW 01/2016!!! Purges all L5000 (FI match rec'd, not eligible for levy) worklist items from your USWT."
  Text 130, 200, 290, 30, "-- NEW 01/2016!!! Reviews M6529 (review for continued interest suspension) worklist items, presenting the information related and giving the worker the choice of whether or not to purge the worklist item."
  Text 110, 180, 280, 10, "-- NEW 01/2016!!! Purges all M8001 (review case referred) worklist items from your USWT."
  Text 100, 230, 340, 10, "-- NEW 01/2016!!! Reviews all L2500 and L2501 (quarterly wage info for CP and NCP) from your USWT.  "
  Text 120, 250, 340, 10, "-- NEW 02/2016!!! Scrubs E4111 (review payment plan) worklists when DL is already suspended."
  Text 140, 270, 300, 20, "-- NEW 02/2016!!! Presents recent payment activity to the user to evaluate E4111 (review pay plan) worklists."
EndDialog



'THE SCRIPT-----------------------------------------------------------------------------------------------
'Shows the dialog
DO
	Dialog BULK_main_menu_dialog
	If buttonpressed = cancel then StopScript
	IF ButtonPressed = SIR_button THEN CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/PRISMscripts/PRISM%20script%20wiki/Forms/AllPages.aspx")
LOOP UNTIL ButtonPressed <> SIR_button
IF ButtonPressed = BULK_cali_to_excel_button 					THEN CALL run_from_GitHub(script_repository & "BULK/BULK - CALI TO EXCEL.vbs")
IF ButtonPressed = BULK_case_transfer_button 					THEN CALL run_from_GitHub(script_repository & "BULK/BULK - CASE TRANSFER.vbs")
IF ButtonPressed = BULK_cp_companion_case_finder_button 			THEN CALL run_from_GitHub(script_repository & "BULK/BULK - CP COMPANION CASE FINDER.vbs")
IF ButtonPressed = BULK_ncp_companion_case_finder_button			THEN CALL run_from_GitHub(script_repository & "BULK/BULK - NCP COMPANION CASE FINDER.vbs")
IF ButtonPressed = BULK_failure_pof_rsdi_dfas_button				THEN CALL run_from_GitHub(script_repository & "BULK/BULK - FAILURE POF RSDI DFAS.vbs")
IF ButtonPressed = BULK_evaluate_nocs_button 					THEN CALL run_from_GitHub(script_repository & "BULK/BULK - EVALUATE NOCS.vbs")
IF ButtonPressed = BULK_L5000_button						THEN CALL run_from_GitHub(script_repository & "BULK/BULK - L5000 WORKLIST SCRUBBER.vbs")
IF ButtonPressed = BULK_M6529_button						THEN CALL run_from_GitHub(script_repository & "BULK/BULK - M6529.vbs")
IF ButtonPressed = BULK_M8001_button						THEN CALL run_from_GitHub(script_repository & "BULK/BULK - M8001 WORKLIST SCRUBBER.vbs")
IF ButtonPressed = BULK_REVIEW_QW_button					THEN CALL run_from_GitHub(script_repository & "BULK/BULK - REVIEW QW INFO.vbs")
IF ButtonPressed = BULK_E4111_activity_button				THEN CALL run_from_GitHub(script_repository & "BULK/BULK - E4111 WORKLIST SCRUBBER.vbs")
IF ButtonPressed = BULK_E4111_suspended_button				THEN CALL run_from_GitHub(script_repository & "BULK/BULK - E4111 SUSP SCRUBBER.vbs")


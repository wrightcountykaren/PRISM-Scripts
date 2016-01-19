'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MAIN MENU.vbs"
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
BeginDialog BULK_main_menu_dialog, 0, 0, 381, 275, "BULK Main Menu"
  ButtonGroup ButtonPressed
    PushButton 5, 30, 60, 10, "CALI to Excel", BULK_cali_to_excel_button
    PushButton 5, 55, 60, 10, "Case Transfer", BULK_case_transfer_button
    PushButton 5, 80, 100, 10, "Companion Case Finder - CP", BULK_cp_companion_case_finder_button
    PushButton 5, 100, 105, 10, "Companion Case Finder - NCP", BULK_ncp_companion_case_finder_button
    PushButton 5, 120, 60, 10, "Evaluate NOCS", BULK_evaluate_nocs_button
    PushButton 5, 140, 90, 10, "Failure POF -- SSA, DFAS", BULK_failure_pof_rsdi_dfas_button
    PushButton 5, 160, 85, 10, "F0320 Worklist Scrubber", BULK_F0320_button
    PushButton 5, 180, 85, 10, "L5000 Worklist Scrubber", BULK_L5000_button
    PushButton 5, 195, 85, 10, "M6529 Worklist Scrubber", BULK_M6529_button
    PushButton 5, 215, 85, 10, "M8001 Worklist Scrubber", BULK_M8001_button
    PushButton 5, 230, 100, 10, "Review Quarterly Wage Info", BULK_REVIEW_QW_button
    CancelButton 325, 255, 50, 15
    PushButton 300, 5, 75, 10, "PRISM Scripts in SIR", SIR_button
  Text 75, 30, 295, 20, "-- This script builds a list in Microsoft Excel of case numbers, function types, program codes, interstate codes, and parent names."
  Text 75, 55, 295, 20, "-- This script allows users to transfer up to 15 cases to as many workers as they need OR to transfer an entire caseload to as many workers as needed."
  Text 110, 80, 260, 10, "--- This script builds a list of companion cases for your CPs on a given CALI."
  Text 115, 100, 255, 10, "-- This script builds a list of companion cases for your NCPs on a given CALI."
  Text 75, 120, 260, 10, "-- This script evaluates D0800 worklist items for continued services."
  Text 100, 140, 270, 15, "-- Clears E0014 worklist item when income is from RSDI (US Treasury) or Dept of Defense."
  Text 95, 160, 280, 20, "-- NEW 01/2016!!! Purges F0320 worklist items when the Med Code is "PAO" and the order type is not "CTM.""
  Text 95, 180, 280, 10, "-- NEW 01/2016!!! Purges all L5000 worklist items from your USWT."
  Text 95, 195, 280, 20, "-- NEW 01/2016!!! Reviews M6529 worklist items, presenting the information related and giving the worker the choice of whether or not to purge the worklist item."
  Text 95, 215, 280, 10, "-- NEW 01/2016!!! Purges all M8001 worklist items from your USWT."
  Text 110, 230, 260, 10, "-- NEW 01/2016!!! Reviews all quarterly wage info from your USWT."
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
IF ButtonPressed = BULK_F0320_button 						THEN CALL run_from_GitHub(script_repository & "BULK/BULK - F0320 SCRUBBER.vbs")
IF ButtonPressed = BULK_L5000_button						THEN CALL run_from_GitHub(script_repository & "BULK/BULK - L5000 WORKLIST SCRUBBER.vbs")
IF ButtonPressed = BULK_M6529_button						THEN CALL run_from_GitHub(script_repository & "BULK/BULK - M6529.vbs")
IF ButtonPressed = BULK_M8001_button						THEN CALL run_from_GitHub(script_repository & "BULK/BULK - M8001 WORKLIST SCRUBBER.vbs")
IF ButtonPressed = BULK_REVIEW_QW_button					THEN CALL run_from_GitHub(script_repository & "BULK/BULK - REVIEW QW INFO.vbs")

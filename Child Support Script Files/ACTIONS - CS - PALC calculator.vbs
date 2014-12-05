'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - CS - PALC calculator"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS---------------------------------------------------------------------------
BeginDialog start_end_date_dialog, 0, 0, 171, 65, "Start and End Date Dialog"
  ButtonGroup ButtonPressed
    OkButton 120, 25, 50, 15
    CancelButton 120, 45, 50, 15
  EditBox 60, 5, 110, 15, PRISM_case_number
  EditBox 45, 25, 70, 15, start_date
  EditBox 45, 45, 70, 15, end_date
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 40, 10, "Start date:"
  Text 5, 50, 40, 10, "End date:"
EndDialog


'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

call PRISM_case_number_finder(PRISM_case_number)

Do
	Do
		Dialog start_end_date_dialog				'Shows dialog
		If buttonpressed = 0 then stopscript		'Cancel
		call PRISM_case_number_validation(PRISM_case_number, case_number_is_valid)
		If case_number_is_valid = False then MsgBox "Your case number isn't valid. Try again."
	Loop until case_number_is_valid = True
	If isdate(start_date) = False or isdate(end_date) = False then MsgBox "You must enter valid dates for both the start and end dates."		'Because a date for both fields is required
Loop until isdate(start_date) = True and isdate(end_date) = True



'Checks to make sure PRISM isn't locked out
transmit
PRISM_check_function

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to PALC
call navigate_to_PRISM_screen("PALC")

'Entering case number and transmitting
EMSetCursor 20, 9
EMSendKey replace(PRISM_case_number, "-", "")	 	'Entering the specific case indicated
transmit								'Transmitting into it



row = 9		'Setting variable for the do...loop

Do

	'Reading payment date, which for some crazy reason is YYMMDD, without slashes. This converts.
	EMReadScreen pmt_ID_YY, 2, row, 7
	EMReadScreen pmt_ID_MM, 2, row, 9
	EMReadScreen pmt_ID_DD, 2, row, 11
	pmt_ID_date = pmt_ID_MM & "/" & pmt_ID_DD & "/" & pmt_ID_YY	

		If (cdate(start_date) <= cdate(pmt_ID_date)) and (cdate(pmt_ID_date) <= cdate(end_date)) then 				'Checks to see if date is in between start/end dates
		date_within_range = True															'Determines date range
		EMReadScreen proc_type, 3, row, 25														'Reading the proc type
		EMReadScreen case_alloc_amt, 10, row, 70													'Reading the amt allocated
		If proc_type = "FTS" or proc_type = "MCE" or proc_type = "NOC" or proc_type = "IFC" or proc_type = "OST" or _	
		proc_type = "PCA" or proc_type = "PIF" or proc_type = "STJ" or proc_type = "STS" or proc_type = "FTJ" then 		'If proc type is one of these, it's involuntary. Else, it's voluntary.
			total_involuntary_alloc = total_involuntary_alloc + abs(case_alloc_amt)							'Adds the alloc amt for involuntary
		Else
			total_voluntary_alloc = total_voluntary_alloc + abs(case_alloc_amt)							'Adds the alloc amt for voluntary
		End if
	End if
	row = row + 1														'Increases the row variable by one, to check the next row
	EMReadScreen end_of_data_check, 19, row, 28									'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do							'Exits do if we have
	If row = 19 then														'Resets row and PF8s
		PF8
		row = 9
	End if
Loop until end_of_data_check = "*** End of Data ***"

If total_involuntary_alloc = "" then total_involuntary_alloc = "0"
If total_voluntary_alloc = "" then total_voluntary_alloc = "0"

string_for_msgbox = "---PAYMENT BREAKDOWN FOR " & start_date & " THROUGH " & end_date & "---" & chr(10) & chr(10) & "Involuntary: $" & total_involuntary_alloc & chr(10) & "Voluntary: $" & total_voluntary_alloc

MsgBox string_for_msgbox

script_end_procedure("")

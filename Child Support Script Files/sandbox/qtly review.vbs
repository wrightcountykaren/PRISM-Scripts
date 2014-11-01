'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = ""
start_time = timer

''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'<<<<<PRISM SPECIFIC, MERGE INTO MAIN FUNCTIONS FILE BEFORE GO-LIVE
Function PRISM_case_number_validation(case_number_to_validate, outcome)
  If len(case_number_to_validate) <> 13 then 
    outcome = False
  Elseif isnumeric(left(case_number_to_validate, 10)) = False then
    outcome = False
  Elseif isnumeric(right(case_number_to_validate, 2)) = False then
    outcome = False
  Elseif InStr(11, case_number_to_validate, "-") <> 11 then
    outcome = False
  Else
    outcome = True
  End if
End function

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
days_to_push_out_worklist = "90"	'This is the default

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog quarterly_reviews_dialog, 0, 0, 176, 85, "Quarterly Reviews Dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog

BeginDialog payment_status_dialog, 0, 0, 231, 100, "payment status dialog"
  DropListBox 125, 25, 80, 15, " "+chr(9)+"NCP is paying"+chr(9)+"NCP is not paying", NCP_payment_status
  ButtonGroup ButtonPressed
    OkButton 50, 60, 50, 15
    CancelButton 135, 60, 50, 15
  Text 15, 25, 95, 15, "Payment Status (select one)"
EndDialog





'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Sends a transmit to check for password issues
transmit

'Checking to make sure we're on USWT or USWD. If not the script will stop.
EMReadScreen USW_check, 3, 21, 75
If USW_check <> "USW" then script_end_procedure("USWT or USWD screen not found. Please start this script from the worklist you are trying to copy over.")

'Searches for the case number.
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if

'<<<<A TEMPORARY MSGBOX TO CHECK THE ACCURACY OF THE PRISM CASE NUMBER FINDER. IF THIS WORKS CREATE A CUSTOM FUNCTION OUT OF THE ABOVE CODE
If PRISM_case_number <> "" then MsgBox "A case number was automatically found on this screen! It is indicated as: " & PRISM_case_number & ". If this case number is incorrect, please take a screenshot of PRISM and send a description of what's wrong to Veronica Cary."

Do
	Do
		Do
			dialog quarterly_reviews_dialog
			If buttonpressed = 0 then stopscript
			call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
			If case_number_valid = False then MsgBox("Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''")
		Loop until case_number_valid = True
		If isnumeric(days_to_push_out_worklist) = False then MsgBox ("You must put a number in for the days to push out worklist.")
	Loop until isnumeric(days_to_push_out_worklist) = True


	EMReadScreen worklist_line_01, 72, 10, 4			'Reads worklist info, line by line
	EMReadScreen worklist_line_02, 72, 11, 4
	EMReadScreen worklist_line_03, 72, 12, 4
	EMReadScreen worklist_line_04, 72, 13, 4
	EMWriteScreen "__________", 17, 21				'clearing out worklist date
	EMWriteScreen days_to_push_out_worklist, 17, 52		'Adding the number of days to push out worklist
	EMWriteScreen "m", 3, 30					'Must modify the panel
	transmit
	call navigate_to_PRISM_screen("CAAD")
	pf5
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then MsgBox "The script could not navigate to a case note. You might be locked out of your case. Navigate to a blank case note and try again."
Loop until case_activity_detail = "Case Activity Detail"

EMWriteScreen worklist_line_01, 16, 4	
EMWriteScreen worklist_line_02, 17, 4
EMWriteScreen worklist_line_03, 18, 4
EMWriteScreen worklist_line_04, 19, 4
EMWriteScreen "------" & worker_signature, 20, 4
EMWriteScreen "E0002", 4, 54



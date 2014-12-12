'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BETA - NOTE - CS - order"
start_time = timer

'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'FUNCTIONS (MERGE INTO MAIN FUNCTIONS BEFORE GO-LIVE)----------------------------------------------------------------------------------------------------
Function convert_CO_FIPS_to_state(state_code, variable)
  If state_code = "01" then variable = "ALABAMA"
  If state_code = "02" then variable = "ALASKA"
  If state_code = "04" then variable = "ARIZONA"
  If state_code = "05" then variable = "ARKANSAS"
  If state_code = "06" then variable = "CALIFORNIA"
  If state_code = "08" then variable = "COLORADO"
  If state_code = "09" then variable = "CONNECTICUT"
  If state_code = "10" then variable = "DELAWARE"
  If state_code = "11" then variable = "DC"
  If state_code = "12" then variable = "FLORIDA"
  If state_code = "13" then variable = "GEORGIA"
  If state_code = "15" then variable = "HAWAII"
  If state_code = "16" then variable = "IDAHO"
  If state_code = "17" then variable = "ILLINOIS"
  If state_code = "18" then variable = "INDIANA"
  If state_code = "19" then variable = "IOWA" 
  If state_code = "20" then variable = "KANSAS"
  If state_code = "21" then variable = "KENTUCKY"
  If state_code = "22" then variable = "LOUISIANA"
  If state_code = "23" then variable = "MAINE"
  If state_code = "24" then variable = "MARYLAND"     
  If state_code = "25" then variable = "MASSACHUSETTS"
  If state_code = "26" then variable = "MICHIGAN"     
  If state_code = "27" then variable = "MINNESOTA"    
  If state_code = "28" then variable = "MISSISSIPPI"  
  If state_code = "29" then variable = "MISSOURI"     
  If state_code = "30" then variable = "MONTANA"      
  If state_code = "31" then variable = "NEBRASKA"     
  If state_code = "32" then variable = "NEVADA"       
  If state_code = "33" then variable = "NEW HAMPSHIRE"
  If state_code = "34" then variable = "NEW JERSEY"    
  If state_code = "35" then variable = "NEW MEXICO"    
  If state_code = "36" then variable = "NEW YORK"      
  If state_code = "37" then variable = "NORTH CAROLINA"
  If state_code = "38" then variable = "NORTH DAKOTA"  
  If state_code = "39" then variable = "OHIO"
  If state_code = "40" then variable = "OKLAHOMA"
  If state_code = "41" then variable = "OREGON"
  If state_code = "42" then variable = "PENNSYLVANIA"
  If state_code = "44" then variable = "RHODE ISLAND"
  If state_code = "45" then variable = "SOUTH CAROLINA"
  If state_code = "46" then variable = "SOUTH DAKOTA"
  If state_code = "47" then variable = "TENNESSEE"
  If state_code = "48" then variable = "TEXAS"
  If state_code = "49" then variable = "UTAH"
  If state_code = "50" then variable = "VERMONT"      
  If state_code = "51" then variable = "VIRGINIA"     
  If state_code = "53" then variable = "WASHINGTON"   
  If state_code = "54" then variable = "WEST VIRGINIA"
  If state_code = "55" then variable = "WISCONSIN"    
  If state_code = "56" then variable = "WYOMING"       
  If state_code = "66" then variable = "GUAM"          
  If state_code = "72" then variable = "PUERTO RICO"   
  If state_code = "78" then variable = "VIRGIN ISLANDS"
  If state_code = "8A" then variable = "SASKATCHEWAN"  
  If state_code = "8B" then variable = "NW TERRITORIES" 
  If state_code = "8C" then variable = "NUNAVUT"        
  If state_code = "8D" then variable = "YUKON TERRITORY"
  If state_code = "80" then variable = "INTERNATIONAL"  
  If state_code = "81" then variable = "ALBERTA"        
  If state_code = "82" then variable = "BRITISH COLUMBIA"
  If state_code = "83" then variable = "MANITOBA"        
  If state_code = "84" then variable = "NEW BRUNSWICK"
  If state_code = "85" then variable = "NEW FOUNDLAND"
  If state_code = "86" then variable = "NOVA SCOTIA"
  If state_code = "87" then variable = "ONTARIO"
  If state_code = "88" then variable = "PRINCE EDWARD ISLAND"
  If state_code = "89" then variable = "QUEBEC"
  If state_code = "90" then variable = "TRIBAL NATIONS"
  If state_code = "93" then variable = "ARMED FORCES, NOT CA"
  If state_code = "94" then variable = "ARMED FORCES, EUROPE"
  If state_code = "95" then variable = "ARMED FORCES, PACIF"

End function


'Function PRISM_case_number_validation(case_number_to_validate, outcome)
'  If len(case_number_to_validate) <> 13 then 
'    outcome = False
'  Elseif isnumeric(left(case_number_to_validate, 10)) = False then
'    outcome = False
'  Elseif isnumeric(right(case_number_to_validate, 2)) = False then
'    outcome = False
'  Elseif InStr(11, case_number_to_validate, "-") <> 11 then
'    outcome = False
'  Else
'    outcome = True
'  End if
'End function



'Function write_editbox_in_PRISM_case_note(x, y, z) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
'  EMGetCursor row, col 
'  EMReadScreen line_check, 2, 15, 2
'  If ((row = 20 and col + (len(x)) >= 78) or row = 21) and line_check = "26" then 
'    MsgBox "You've run out of room in this case note. The script will now stop."
'    StopScript
'  End if
'  If row = 21 then
'    EMSendKey "<PF8>"
'    EMWaitReady 0, 0
'    EMSetCursor 16, 4
'  End if
'  variable_array = split(y, " ")
'  EMSendKey "* " & x & ": "
'  For each x in variable_array 
'    EMGetCursor row, col 
'    EMReadScreen line_check, 2, 15, 2
'    If ((row = 20 and col + (len(x)) >= 78) or row = 21) and line_check = "26" then 
'      MsgBox "You've run out of room in this case note. The script will now stop."
'      StopScript
'    End if
'    If (row = 20 and col + (len(x)) >= 78) or (row = 16 and col = 4) or row = 21 then
'      EMSendKey "<PF8>"
'      EMWaitReady 0, 0
'      EMSetCursor 16, 4
'    End if
'    EMGetCursor row, col 
'    If (row < 20 and col + (len(x)) >= 78) then EMSendKey "<newline>" & space(z)
'    If (row = 16 and col = 4) then EMSendKey space(z)
'    EMSendKey x & " "
'    If right(x, 1) = ";" then 
'      EMSendKey "<backspace>" & "<backspace>" 
'      EMGetCursor row, col 
'      If row = 20 then
'        EMSendKey "<PF8>"
'        EMWaitReady 0, 0
'        EMSetCursor 16, 4
'        EMSendKey space(z)
'      Else
'        EMSendKey "<newline>" & space(z)
'      End if
'    End if
'  Next
'  EMSendKey "<newline>"
'  EMGetCursor row, col 
'  If (row = 20 and col + (len(x)) >= 78) or (row = 16 and col = 4) then
'    EMSendKey "<PF8>"
'    EMWaitReady 0, 0
'    EMSetCursor 16, 4
'  End if
'End function

'Function write_new_line_in_PRISM_case_note(x)
'  EMGetCursor row, col 
'  EMReadScreen line_check, 2, 15, 2
'  If ((row = 20 and col + (len(x)) >= 78) or row = 21) and line_check = "26" then 
'    MsgBox "You've run out of room in this case note. The script will now stop."
'    StopScript
'  End if
'  If (row = 20 and col + (len(x)) >= 78 + 1 ) or row = 21 then
'    EMSendKey "<PF8>"
'    EMWaitReady 0, 0
'    EMSetCursor 16, 4
'  End if
'  EMSendKey x & "<newline>"
'  EMGetCursor row, col 
'  If (row = 20 and col + (len(x)) >= 78) or (row = 21) then
'    EMSendKey "<PF8>"
'    EMWaitReady 0, 0
'    EMSetCursor 16, 4
'  End if
'End function

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 70, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  EditBox 85, 30, 75, 15, date_of_order
  ButtonGroup ButtonPressed
    OkButton 35, 50, 50, 15
    CancelButton 95, 50, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
  Text 30, 35, 50, 10, "Date of order:"
EndDialog

'NOTE: I had to manually add the DropListBox options because the dialog editor could not handle the whole list
BeginDialog order_dialog, 0, 0, 386, 260, "Order dialog"
  DropListBox 60, 5, 220, 15, "select one"+chr(9)+"COMPELLED TO GENETIC TESTING"+chr(9)+"CONTINUANCE ORDER"+chr(9)+"CONTEMPT/ADJ PLUS SUPPORT MODI"+chr(9)+"CONTEMPT"+chr(9)+"CONTEMPT/POF"+chr(9)+"DOMESTIC ABUSE"+chr(9)+"DISMISSAL-UIFSA-PAT"+chr(9)+"DISMISSAL-CONTEMPT"+chr(9)+"DISMISSAL-ESTABLISH"+chr(9)+"DISMISSAL-MODIFICATION"+chr(9)+"DISMISSAL-PATERNITY"+chr(9)+"DISSOLUTION"+chr(9)+"ENFORCE OTHER STATE ORDER"+chr(9)+"LICENSE SUSPENSION APPEAL"+chr(9)+"COLA - MOTION DENIED"+chr(9)+"MODIFICATION - MOTION DENIED"+chr(9)+"MINNESOTA STATUTE 256"+chr(9)+"SUPPORT SET BASED ON A MOTION"+chr(9)+"PATERNITY"+chr(9)+"REVIEW ADJUST/MODIFICATION"+chr(9)+"REDIRECTION OF SUPPORT"+chr(9)+"SUPPORT ESTABLISHED"+chr(9)+"TEMPORARY"+chr(9)+"TERMINATION OF PARENTAL RIGHTS"+chr(9)+"TEMPORARY PATERNITY"+chr(9)+"UIFSA - CONTROLLING ORDER"+chr(9)+"UIFSA - ESTABLISH"+chr(9)+"UIFSA - REGISTER FOR MOD"+chr(9)+"UIFSA - REGISTER FOR ENFC"+chr(9)+"UIFSA - PATERNITY"+chr(9)+"VACATE - ESTABLISH"+chr(9)+"VACATE - PATERNITY", order_header
  EditBox 85, 25, 295, 15, NCOD_info
  EditBox 85, 45, 295, 15, eff_date_of_obligation
  CheckBox 20, 80, 25, 10, "GR?", emancipation_GR_check
  CheckBox 65, 80, 25, 10, "DS?", emancipation_DS_check
  CheckBox 110, 80, 25, 10, "OT?", emancipation_OT_check
  CheckBox 155, 80, 25, 10, "18?", emancipation_18_check
  CheckBox 200, 80, 25, 10, "21?", emancipation_21_check
  CheckBox 20, 95, 195, 10, "Check here if there is a reduction due to emancipation.", reduction_check
  EditBox 50, 110, 325, 15, emancipation_language_notes
  EditBox 35, 135, 100, 15, un_un
  EditBox 205, 145, 25, 15, who_is_required_to_carry_medical
  EditBox 205, 165, 25, 15, who_is_required_to_carry_dental
  EditBox 295, 145, 25, 15, who_is_required_to_carry_med_support
  EditBox 120, 200, 260, 15, review_hearings_and_conditions
  EditBox 95, 220, 120, 15, whose_order
  EditBox 70, 240, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 275, 240, 50, 15
    CancelButton 330, 240, 50, 15
  Text 5, 10, 50, 10, "CAAD header:"
  Text 5, 30, 75, 10, "NCOD info (type, amt):"
  Text 5, 50, 70, 10, "Eff date of obligation:"
  GroupBox 5, 70, 375, 60, "Emancipation language"
  Text 20, 115, 25, 10, "Notes:"
  Text 5, 140, 25, 10, "Un/Un:"
  GroupBox 165, 135, 160, 55, "Who is required to carry..."
  Text 175, 150, 30, 10, "Medical:"
  Text 175, 170, 25, 10, "Dental:"
  Text 245, 150, 45, 10, "Med Support:"
  Text 5, 205, 110, 10, "Review hearings and conditions:"
  Text 5, 225, 90, 10, "Whose order (state-FIPS):"
  Text 5, 245, 65, 10, "Worker Signature:"
EndDialog







'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

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

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		If buttonpressed = 0 then stopscript
		call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	Loop until case_number_valid = True
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"



'Autofilling from PRISM------------------------

'Clearing regional globals
call navigate_to_PRISM_screen("REGL")
transmit

'Getting info from SUOD
call navigate_to_PRISM_screen("SUOD")
EMSetCursor 4, 8														'Setting the cursor on the case line
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "N", 3, 29												'Setting the screen as a display action
transmit															'Transmitting into it
EMReadScreen FIPS_state_code, 2, 4, 55										'Reading state code
call convert_CO_FIPS_to_state(FIPS_state_code, whose_order)							'Converting state code to a state name
PF11																'Going to the next screen
EMReadScreen un_un_NCP, 3, 12, 37											'Reading NCP un/un expense
EMReadScreen un_un_CP, 3, 12, 48											'Reading CP un/un expense
If un_un_CP = "___" then												
	un_un_CP = "N/A"
Else
	un_un_CP = cint(un_un_CP)& "%"
End if
If un_un_NCP = "___" then 
	un_un_NCP = "N/A"
Else
	un_un_NCP = cint(un_un_NCP)& "%"
End if
un_un = "NCP- " & un_un_NCP & ", CP- " & un_un_CP  								'Converting the previous two variables into one
EMReadScreen who_is_required_to_carry_medical, 3, 10, 48							'Reading who_is_required_to_carry_medical
who_is_required_to_carry_medical = replace(who_is_required_to_carry_medical, "_", "")		'Cleaning up the variable we just read
EMReadScreen who_is_required_to_carry_dental, 3, 11, 48							'Reading who_is_required_to_carry_dental
who_is_required_to_carry_dental = replace(who_is_required_to_carry_dental, "_", "")			'Cleaning up the variable we just read
EMReadScreen who_is_required_to_carry_med_support, 3, 12, 74						'Reading who_is_required_to_carry_med_support
who_is_required_to_carry_med_support = replace(who_is_required_to_carry_med_support, "_", "")	'Cleaning up the variable we just read

'Getting info from NCOL
PRISM_row = 9 															'Setting the variable for the following do...loop
Do
	call navigate_to_PRISM_screen("NCOL")

	EMWriteScreen "___", 20, 39			'Removing the different default obl type
	EMWriteScreen "12/31/2099", 20, 58		'Changing the default date so we can see the future
	EMWriteScreen "A", 20, 77			'We only want to see active claims
	transmit

	EMReadScreen line_check, 2, PRISM_row, 5	'checking to see if there's anything on this row.
	If line_check = "  " then exit do			'If nothing is on this row, the do...loop will end.
	EMWriteScreen "S", PRISM_row, 3			'selecting a row to evaluate.
	transmit


	'<<<<<<<<<<<<NOTE: commented out date-related logic. If this works, remove from script before going live. 04/17/2014 
	EMReadScreen obligation_type, 3, 4, 34
	'EMReadScreen effective_date, 7, 9, 59
	'effective_date_comparison_variable = cdate(left(effective_date, 2) & "/01/" & (right(effective_date, 2)))
	EMReadScreen monthly_accrual, 9, 16, 19
	monthly_accrual = cint(trim(monthly_accrual))
	'If datediff("m", effective_date_comparison_variable, date_of_order) = 0 then
		'NCOD_info = NCOD_info & obligation_type & " ($" & monthly_accrual & " eff " & effective_date & "), "
	'Elseif datediff("m", effective_date_comparison_variable, date_of_order) < 0 then
		'NCOD_info_future = NCOD_info_future & obligation_type & " ($" & monthly_accrual & " eff " & effective_date & "), "
	'End if
	NCOD_info = NCOD_info & obligation_type & " ($" & monthly_accrual & "), "	'<<<<<<<<<Added this line 04/17/2014, see above comments
	PRISM_row = PRISM_row + 1
Loop until line_check = "  "


'Cleaning up NCOD info (removing end commas)
If right(NCOD_info, 2) = ", " then NCOD_info = left(NCOD_info, len(NCOD_info) - 2)
'If right(NCOD_info_future, 2) = ", " then NCOD_info_future = left(NCOD_info_future, len(NCOD_info_future) - 2)	'<<<<<<<<<<REMOVED as this relates to date-logic

'Checks to see if we're on the case activity detail screen. If we aren't, the script will end.
Do
	Do
		dialog order_Dialog
		If buttonpressed = 0 then stopscript
		If order_header = "select one" then MsgBox("You must select an order type, it's required for the header of your case note.")
	Loop until order_header <> "select one"
	call navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then MsgBox "The script could not navigate to a case note. You might be locked out of your case. Navigate to a blank case note and try again."
Loop until case_activity_detail = "Case Activity Detail"

'Combining emancipation language so that it appears as a single line
If emancipation_GR_check = 1 then emancipation_language = emancipation_language & "GR, "
If emancipation_DS_check = 1 then emancipation_language = emancipation_language & "DS, "
If emancipation_OT_check = 1 then emancipation_language = emancipation_language & "OT, "
If emancipation_18_check = 1 then emancipation_language = emancipation_language & "18, "
If emancipation_21_check = 1 then emancipation_language = emancipation_language & "21, "
If right(emancipation_language, 2) = ", " then emancipation_language = left(emancipation_language, len(emancipation_language) - 2)

'Writing the activity date and "free" code to the CAAD note. The activity date is the date of the order.
EMWriteScreen date_of_order, 4, 37
EMWriteScreen "free", 4, 54

'Writing the case note
EMSetCursor 16, 4
EMSendKey "***" & order_header & "***" & "<newline>"
If date_of_order <> "" then call write_editbox_in_PRISM_case_note("Date of order", date_of_order, 5)
If NCOD_info <> "" then call write_editbox_in_PRISM_case_note("NCOD info", NCOD_info, 5)
If eff_date_of_obligation <> "" then call write_editbox_in_PRISM_case_note("Eff Date of Obligation", eff_date_of_obligation, 5)
If emancipation_language <> "" then call write_editbox_in_PRISM_case_note("Emancipation language", emancipation_language, 5)
If reduction_check = 1 then call write_new_line_in_PRISM_case_note("* This was a reduction.")
If emancipation_language_notes <> "" then call write_editbox_in_PRISM_case_note("Notes", emancipation_language_notes, 5)
If un_un <> "" then call write_editbox_in_PRISM_case_note("Un/Un", un_un, 5)
If who_is_required_to_carry_medical <> "" then call write_editbox_in_PRISM_case_note("Who is required to carry medical?", who_is_required_to_carry_medical, 5)
If who_is_required_to_carry_dental <> "" then call write_editbox_in_PRISM_case_note("Who is required to carry dental?", who_is_required_to_carry_dental, 5)
If who_is_required_to_carry_med_support <> "" then call write_editbox_in_PRISM_case_note("Who is required to carry med support?", who_is_required_to_carry_med_support, 5)
If review_hearings_and_conditions <> "" then call write_editbox_in_PRISM_case_note("Review hearings and conditions", review_hearings_and_conditions, 5)
If whose_order <> "" then call write_editbox_in_PRISM_case_note("Whose order?", whose_order, 5)
call write_new_line_in_PRISM_case_note("---")
call write_new_line_in_PRISM_case_note(worker_signature)

script_end_procedure("")
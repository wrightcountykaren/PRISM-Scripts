'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - CS - client contact"
start_time = timer

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
date_of_contact = date & ""	'defaults to today

'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support Script Files\FUNCTIONS FILE.vbs")
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



'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog contact_dialog, 0, 0, 381, 295, "Client contact"
  DropListBox 80, 15, 260, 15, ""+chr(9)+"T0050 PHONE CALL TO CP"+chr(9)+"T0051 PHONE CALL FR CP"+chr(9)+"T0052 PHONE CALL RET TO CP"+chr(9)+"T0053 PHONE CALL RET FR CP"+chr(9)+"T0054 PHONE CALL ATMPT TO RET TO CP"+chr(9)+"T0093 CONTACT WITH CP SPOUSE"+chr(9)+"T0101 PHONE CONTACT CP'S ATTORNEY"+chr(9)+"T0201 CONTACT WITH CP EMPLOYER"+chr(9)+"M3910 INTERVIEW WITH CP", contact_type_CP
  DropListBox 80, 35, 260, 15, ""+chr(9)+"T0055 PHONE CALL TO NCP"+chr(9)+"T0056 PHONE CALL FR NCP"+chr(9)+"T0057 PHONE CALL RET TO NCP"+chr(9)+"T0058 PHONE CALL RET FR NCP"+chr(9)+"T0059 PHONE CALL ATMPT TO RET TO NCP"+chr(9)+"T0060 PHONE CALL TO NCP EMP"+chr(9)+"T0061 PHONE CALL FROM NCP EMP"+chr(9)+"T0062 PHONE CALL RET TO NCP EMP"+chr(9)+"T0063 PHONE CALL RET FR NCP EMP"+chr(9)+"T0064 PHONE CALL ATMPT RET TO NCP EMP"+chr(9)+"T0065 PHONE CALL TO NCP AY"+chr(9)+"T0066 PHONE CALL FR NCP AY"+chr(9)+"T0067 PHONE CALL RET TO NCP AY"+chr(9)+"T0068 PHONE CALL RET FR NCP AY"+chr(9)+"T0069 PHONE CALL ATMPT RET TO NCP AY"+chr(9)+"T0092 CONTACT WITH NCP SPOUSE"+chr(9)+"M3911 INTERVIEW WITH NCP", contact_type_NCP
  DropListBox 80, 55, 260, 15, ""+chr(9)+"M0410 CONTACT WITH CCC WORKER"+chr(9)+"T0070 PHONE CALL/OTHER"+chr(9)+"T0074 CONTACT WITH STATE HELP DESK"+chr(9)+"T0075 CONTACT WITH HEALTH CARRIER"+chr(9)+"T0080 CONTACT WITH COURT ADMINISTRATOR"+chr(9)+"T0085 CONTACT WITH LAW ENFORCEMENT"+chr(9)+"T0087 CONTACT WITH PROBATION OFFICER"+chr(9)+"T0090 CONTACT WITH NCP/CP UNION"+chr(9)+"T0095 CONTACT WITH SOCIAL WORKER"+chr(9)+"T0098 CONTACT WITH WORKER FROM ANOTHER MN COUNTY"+chr(9)+"T0100 PHONE CONTACT WITH OTHER STATE'S CENTRAL REGISTRY"+chr(9)+"T0102 PHONE CONTACT COUNTY ATTORNEY"+chr(9)+"T0103 PHONE CONTACT WITH OTHER STATE WORKER"+chr(9)+"T0104 PHONE CONTACT WITH FINANCIAL WORKER"+chr(9)+"T0105 PHONE CONTACT WITH CSPC"+chr(9)+"T0111 CONTACT WITH OTHER STATE AGENCY", contact_type_other
  EditBox 165, 80, 80, 15, PRISM_case_number
  EditBox 310, 80, 70, 15, date_of_contact
  EditBox 95, 110, 60, 15, phone_number
  EditBox 285, 110, 85, 15, time_contact_was_made
  EditBox 55, 135, 325, 15, issue
  EditBox 55, 155, 325, 15, actions_taken
  EditBox 65, 185, 310, 15, verifs_needed
  EditBox 120, 205, 255, 15, special_instructions_for_client
  CheckBox 5, 230, 150, 10, "Check here if you verified ID.", verified_ID_check
  CheckBox 5, 245, 230, 10, "Check here if you left a generic message requesting they return call.", left_generic_message_check
  EditBox 310, 255, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 270, 275, 50, 15
    CancelButton 330, 275, 50, 15
  Text 15, 20, 55, 10, "CP contact type:"
  Text 15, 40, 60, 10, "NCP contact type:"
  Text 15, 60, 65, 10, "Other contact type:"
  Text 5, 85, 160, 10, "PRISM case number (XXXXXXXXXX-XX format):"
  Text 250, 85, 55, 10, "Date of contact:"
  GroupBox 5, 100, 370, 30, "Optional contact info:"
  Text 40, 115, 50, 10, "Phone number: "
  Text 195, 115, 85, 10, "Time contact was made: "
  Text 5, 140, 50, 10, "Issue/subject: "
  Text 5, 160, 50, 10, "Actions taken: "
  GroupBox 5, 175, 375, 50, "Helpful optional case info"
  Text 15, 190, 50, 10, "Verifs needed: "
  Text 15, 210, 100, 10, "Special instructions for client:"
  Text 235, 260, 70, 10, "Sign your case note: "
  GroupBox 5, 5, 370, 70, "Select one contact type from this group, based on CAAD note requirement"
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
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


'Shows dialog, then navigates to CAAD. It will validate the PRISM case number using the custom function.
Do
	Do
		Do
			Do
				dialog contact_dialog
				If buttonpressed = 0 then stopscript
				call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
				If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
			Loop until case_number_valid = True
			If ((contact_type_CP <> "" and contact_type_NCP = "" and contact_type_other = "") or (contact_type_CP = "" and contact_type_NCP <> "" and contact_type_other = "") or (contact_type_CP = "" and contact_type_NCP = "" and contact_type_other <> "")) = False then MsgBox("Please select one (and only one) of the contact type options.")
		Loop until (contact_type_CP <> "" and contact_type_NCP = "" and contact_type_other = "") or (contact_type_CP = "" and contact_type_NCP <> "" and contact_type_other = "") or (contact_type_CP = "" and contact_type_NCP = "" and contact_type_other <> "")
		If isdate(date_of_contact) = False then MsgBox "You must put a valid date in as the date of contact. Please try again."
	Loop until isdate(date_of_contact) = True
	call navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then MsgBox "The script could not navigate to a case note. You might be locked out of your case. Navigate to a blank case note and try again."
Loop until case_activity_detail = "Case Activity Detail"

'Determining which of the three contact_type editboxes contains info, and then making that the "contact_type" variable
If contact_type_CP <> "" and contact_type_NCP = "" and contact_type_other = "" then contact_type = contact_type_CP
If contact_type_CP = "" and contact_type_NCP <> "" and contact_type_other = "" then contact_type = contact_type_NCP
If contact_type_CP = "" and contact_type_NCP = "" and contact_type_other <> "" then contact_type = contact_type_other


'Writing the case note
EMWriteScreen left(contact_type, 5), 4, 54				'The contact type (only need the left 5 characters)
EMWriteScreen date_of_contact, 4, 37					'Writing the contact date as the activity date on CAAD
EMSetCursor 16, 4 								'Because the PRISM case note functions require the cursor to start here
If issue <> "" then call write_editbox_in_PRISM_case_note("Issue/subject", issue, 5)
If verified_ID_check = 1 then call write_new_line_in_PRISM_case_note("* Verified ID.")
If actions_taken <> "" then call write_editbox_in_PRISM_case_note("Actions taken", actions_taken, 5)
If verifs_needed <> "" then call write_editbox_in_PRISM_case_note("Verifs needed", verifs_needed, 5)
If special_instructions_for_client <> "" then call write_editbox_in_PRISM_case_note("Special Instructions for Client", special_instructions_for_client, 5)
If case_status <> "" then call write_editbox_in_PRISM_case_note("Case status", case_status, 5)
If left_generic_message_check = 1 then call write_new_line_in_PRISM_case_note("* Left client a generic message requesting a return call.")
If phone_number <> "" then call write_editbox_in_PRISM_case_note("Phone number", phone_number, 5)
If time_contact_was_made <> "" then call write_editbox_in_PRISM_case_note("Time contact was made", time_contact_was_made, 5)
call write_new_line_in_PRISM_case_note("---")
call write_new_line_in_PRISM_case_note(worker_signature)


script_end_procedure("")
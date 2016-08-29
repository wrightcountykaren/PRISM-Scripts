'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - DATE OF THE HEARING (EXPRO).vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Using custom functions to convert arrays from global variables into a list for the dialog.
call convert_array_to_droplist_items(county_attorney_array, county_attorney_list)										'County attorneys
call convert_array_to_droplist_items(child_support_magistrates_array, child_support_magistrates_list)					'County judges

BeginDialog date_of_the_hearing_expro_dialog, 0, 0, 321, 220, "Date of the Hearing ExPRO"
  Text 5, 5, 80, 10, "Motion before the Court"
  ComboBox 85, 5, 165, 15, "Select one or type in other motion:"+chr(9)+"MES 256 Action"+chr(9)+"Motion to Set"+chr(9)+"Continuance"+chr(9)+"License Suspension Appeal"+chr(9)+"COLA motion"+chr(9)+"Modification/RAM"+chr(9)+"UFM - Register for Modificaion", motion_before_court
  Text 5, 25, 85, 10, "Child Support Magistrate"
  DropListBox 90, 25, 85, 15, child_support_magistrates_list, child_support_magistrate
  Text 180, 25, 55, 10, "County Attorney"
  DropListBox 235, 25, 85, 15, county_attorney_list, CAO_list
  CheckBox 5, 50, 50, 10, "NCP present", NCP_present_check
  Text 60, 50, 60, 10, "Represented by:"
  EditBox 115, 50, 85, 15, NCP_represented_by
  CheckBox 5, 65, 50, 10, "CP present", CP_present_check
  Text 60, 65, 55, 10, "Represented by:"
  EditBox 115, 65, 85, 15, CP_represented_by
  Text 5, 90, 70, 10, "Details of the hearing"
  EditBox 75, 90, 170, 15, details_of_the_hearing
  CheckBox 5, 110, 100, 10, "Driver's license addressed", DL_addressed_check
  Text 20, 125, 105, 10, "Details of drivers license status"
  EditBox 130, 125, 155, 15, dl_details
  Text 10, 145, 70, 10, "Review Hearing Date"
  EditBox 85, 145, 65, 15, review_hearing_date
  Text 150, 175, 60, 10, "Worker signature"
  EditBox 215, 175, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 200, 50, 15
    CancelButton 255, 200, 50, 15
EndDialog

'case number dialog-
BeginDialog case_number_dialog, 0, 0, 176, 85, "Case number dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog


'Connecting to BlueZone
EMConnect ""

call PRISM_case_number_finder(PRISM_case_number)

'Case number display dialog
Do
	Dialog case_number_dialog
	If buttonpressed = 0 then stopscript
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
Loop until case_number_valid = True



'Displays dialog for date of the hearing caad note and checks for information
Do
	Do
		Do
			Do
				Do
					'Shows dialog, validates that PRISM is up and not timed out, with transmit
					Dialog date_of_the_hearing_expro_dialog
					If buttonpressed = 0 then stopscript
					transmit
					EMReadScreen PRISM_check, 5, 1, 36
					If PRISM_check <> "PRISM" then MsgBox "You appear to have timed out, or are out of PRISM. Navigate to PRISM and try again."
				Loop until PRISM_check = "PRISM"
				'Makes sure worker enters in signature
				If worker_signature = "" then MsgBox "Sign your CAAD note"
			Loop until worker_signature <> ""
			'Makes sure worker selects motion type
			If motion_before_court = "" or motion_before_court = "Select one or type in other motion:" then MsgBox "You must enter in a motion!"
		Loop until motion_before_court <> "" and motion_before_court <> "Select one or type in other motion:"
		'Makes sure worker selects county attorney
		If CAO_list = "Select one:" then MsgBox "Please select a County Attorney"
	Loop until CAO_list <> "Select one:"
	'Makes sure worker selects child support magistrate
	If child_support_magistrate = "Select one:" then MsgBox "Please select a Child Support Magistrate"
Loop until child_support_magistrate <> "Select one:"


'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")

'Entering case number
call enter_PRISM_case_number(PRISM_case_number, 20, 8)


PF5					'Did this because you have to add a new note

EMWriteScreen "M3909", 4, 54  'adds correct caad code

EMSetCursor 16, 4			'Because the cursor does not default to this location

call write_editbox_in_PRISM_case_note("Motion before the Court", motion_before_court, 4)
call write_editbox_in_PRISM_case_note("Child Support Magistrate", child_support_magistrate, 4)
call write_editbox_in_PRISM_case_note("County Attorney", CAO_list, 4)
if NCP_present_check = 1 then
	call write_new_line_in_PRISM_case_note("* NCP present")
	call write_editbox_in_PRISM_case_note("Represented by", NCP_represented_by, 4)
else
	call write_new_line_in_PRISM_case_note ("* NCP not present")
end if
if CP_present_check = 1 then
	call write_new_line_in_PRISM_case_note("* CP present")
	call write_editbox_in_PRISM_case_note("Represented by", CP_represented_by, 4)
else
	call write_new_line_in_PRISM_case_note ("* CP not present")
end if
call write_editbox_in_PRISM_case_note("Details of the Hearing", details_of_the_hearing, 4)
if DL_addressed_check = 1 then
	call write_new_line_in_PRISM_case_note("* Drivers license addressed")
	call write_editbox_in_PRISM_case_note("Details of drivers license", dl_details, 4)
end if
if review_hearing_date <> "" then
	call write_editbox_in_PRISM_case_note("Review Hearing date", review_hearing_date, 4)
end if
call write_new_line_in_PRISM_case_note("---")
call write_new_line_in_PRISM_case_note(worker_signature)

script_end_procedure("")

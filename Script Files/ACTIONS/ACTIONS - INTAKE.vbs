'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - INTAKE.vbs"
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

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
checked = 1
unchecked = 0
CAWD_check = checked
CAAD_note_check = checked

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

BeginDialog CS_intake_dialog, 0, 0, 371, 345, "CS intake dialog"
  EditBox 65, 5, 85, 15, client_last_name
  EditBox 230, 5, 85, 15, client_first_name
  EditBox 60, 35, 120, 15, street_line_1
  EditBox 60, 55, 120, 15, street_line_2
  EditBox 75, 75, 105, 15, city_state_zip
  EditBox 240, 35, 125, 15, NCP_name
  DropListBox 250, 55, 55, 15, "father"+chr(9)+"mother", NCP_gender
  EditBox 250, 75, 115, 15, childs_name
  CheckBox 25, 130, 60, 10, "Child Only MA", child_only_MA_check
  CheckBox 25, 145, 125, 10, "Child Only MA - Relative Caretaker", child_only_MA_relative_caretaker_check
  CheckBox 25, 160, 120, 10, "CP and Child MA", CP_and_child_MA_check
  CheckBox 15, 175, 145, 10, "CP Paternity Request Sheet", CP_paternity_request_sheet_check
  CheckBox 15, 190, 145, 10, "Financial Affidavit OCS", financial_affidavit_OCS_check
  CheckBox 15, 205, 145, 10, "Issues-Paternity-to be Decided", issues_paternity_to_be_decided_check
  CheckBox 15, 220, 145, 10, "Parenting Time Schedules", parenting_time_schedules_check
  CheckBox 25, 250, 60, 10, "Normal", paternity_cover_letter_normal_check
  CheckBox 25, 265, 75, 10, "Relative Caretaker", paternity_cover_letter_relative_caretaker_check
  CheckBox 25, 280, 100, 10, "Minor with GAL Attachment", paternity_cover_letter_minor_check
  CheckBox 15, 295, 145, 10, "Paternity Information Form Memo", paternity_information_form_memo_check
  CheckBox 15, 310, 145, 10, "Paternity Information Form", paternity_information_form_check
  CheckBox 15, 325, 145, 10, "Supplemental Paternity Information Form", supplemental_paternity_information_form_check
  CheckBox 200, 115, 35, 10, "F0018", F0018_check
  CheckBox 200, 130, 35, 10, "F0022", F0022_check
  EditBox 260, 150, 85, 15, worker_name
  EditBox 260, 170, 85, 15, worker_phone
  EditBox 280, 190, 85, 15, worker_signature
  CheckBox 195, 215, 155, 10, "Check here to have script send a CAAD note.", CAAD_note_check
  CheckBox 195, 230, 155, 10, "Check here to have script send a CAWD.", CAWD_check
  ButtonGroup ButtonPressed
    OkButton 230, 250, 50, 15
    CancelButton 285, 250, 50, 15
  Text 5, 10, 55, 10, "Client last name:"
  Text 170, 10, 55, 10, "Client first name:"
  GroupBox 5, 25, 180, 70, "Address"
  GroupBox 190, 100, 180, 45, "DORD docs to print for client"
  Text 205, 195, 70, 10, "Sign your CAAD note:"
  GroupBox 190, 25, 180, 70, "Familial info:"
  Text 10, 60, 45, 10, "Street (line 2):"
  Text 200, 40, 40, 10, "NCP name:"
  Text 10, 80, 65, 10, "City, state and zip:"
  GroupBox 5, 100, 180, 240, "Word docs to print for client"
  Text 15, 115, 80, 10, "Choice of service letter:"
  Text 15, 235, 100, 10, "Paternity Cover Letter to CP:"
  Text 200, 80, 45, 10, "Child's name:"
  Text 10, 40, 45, 10, "Street (line 1):"
  Text 200, 60, 45, 10, "NCP gender:"
  Text 205, 175, 50, 10, "Worker phone:"
  Text 205, 155, 50, 10, "Worker name:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
IF county_name <> "Anoka County" THEN MsgBox "This script contains links to documents stored on the Anoka County network. As such, it may not work for your agency."

'Connects to BlueZone
EMConnect ""

'Finds the PRISM case number using a custom function
call PRISM_case_number_finder(PRISM_case_number)

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

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to CAPS
call navigate_to_PRISM_screen("CAPS")

'Entering case number and transmitting
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit															'Transmitting into it

'Reading info from CAPS
row = 6									'Setting these for an EMSearch for the comma, which signifies the end of a first name. Needs to start looking on row 6.
col = 1									'The col is where the search string must end, need to set as a 1 to show change (it shows a 0 for not found, signfying an error)
EMSearch ", ", row, col							'Finds that comma
last_name_length = col - 12						'12 being the starting column of the CP name
EMReadScreen client_last_name, last_name_length, 6, 12	'Reads the last name based on the length found above
EMReadScreen client_first_name, 20, 6, col + 2			'20 is a nice long amount, and the col variable contains the comma, we need to start reading two columns after that variable
client_first_name = trim(client_first_name)			'Getting rid of the excess spaces
EMReadScreen NCP_name, 40, 7, 12					'This does not get split into separate first/last info, so we'll just read 40 characters and trim
NCP_name = trim(NCP_name)						'Trimming excess spaces
EMReadScreen NCP_PRISM_gender,  3, 18, 76 			'Reading the relationship (FAT if it's father)
If NCP_PRISM_gender = "MOT" then					'If the NCP is "MOT" than NCP should read mother, otherwise it's reversed. It's almost always reversed.
	NCP_gender = "mother"
Else
	NCP_gender = "father"
End if
EMReadScreen childs_name, 30, 18, 16				'This does not get split into separate first/last info, so we'll just read 30 characters and trim
childs_name = trim(childs_name)					'Trimming excess spaces

'Getting all child/DOB info
PRISM_row = 18
Do
	EMReadScreen child_name_on_CAPS, 30, PRISM_row, 16	'reading name
	child_name_on_CAPS = trim(child_name_on_CAPS)		'removing spaces from beginning and end
	EMReadScreen child_DOB_on_CAPS, 10, PRISM_row, 64	'reading DOB
	If child_name_on_CAPS <> "" then CAPS_kids = CAPS_kids & child_name_on_CAPS & " (DOB: " & child_DOB_on_CAPS & ")" & chr(13) 		'If there's a name, add to the CAPS_kids variable
	PRISM_row = PRISM_row + 1					'increase the PRISM row
	If PRISM_row = 21 then						'If we're on row 21, go to the next page
		PF8
		PRISM_row = 18
	End if
Loop until child_name_on_CAPS = ""

'Lower case-ing the intake names
call fix_case(client_first_name, 1)	
call fix_case(client_last_name, 1)
call fix_case(NCP_name, 1)
call fix_case(childs_name, 1)

'Getting worker info for case note
EMSetCursor 5, 53
PF1
EMReadScreen worker_name, 27, 6, 50
EMReadScreen worker_phone, 12, 8, 35
PF3

'Cleaning up worker info
worker_name = trim(worker_name)
call fix_case(worker_name, 1)

'Navigating to CPDD to pull address info
call navigate_to_PRISM_screen("CPDD")
EMReadScreen street_line_1, 30, 15, 11
EMReadScreen street_line_2, 30, 16, 11
EMReadScreen city_state_zip, 49, 17, 11

'Cleaning up address info
street_line_1 = replace(street_line_1, "_", "")
call fix_case(street_line_1, 1)
street_line_2 = replace(street_line_2, "_", "")
call fix_case(street_line_2, 1)
city_state_zip = replace(replace(replace(city_state_zip, "_", ""), "    St: ", ", "), "    Zip: ", " ")
call fix_case(city_state_zip, 2)

'Shows intake dialog, checks to make sure we're still in PRISM (not passworded out)
Do
	Dialog CS_intake_dialog
	If buttonpressed = 0 then stopscript
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"

'Combining variables from dialog
If street_line_2 <> "" then							'Address
	street_address = street_line_1 & chr(13) & street_line_2
Else
	street_address = street_line_1
End if
CP_name = trim(client_last_name) & ", " & trim(client_first_name)			'CP name

'Creating the Word application object (if any of the Word options are selected), and making it visible 
If _
	child_only_MA_check = checked or _
	child_only_MA_relative_caretaker_check = checked or _
	CP_and_child_MA_check = checked or _
	CP_paternity_request_sheet_check = checked or _
	financial_affidavit_OCS_check = checked or _
	paternity_cover_letter_normal_check = checked or _
	paternity_cover_letter_relative_caretaker_check = checked or _
	paternity_cover_letter_minor_check = checked or _
	paternity_information_form_memo_check = checked or _
	paternity_information_form_check = checked or _
	supplemental_paternity_information_form_check = checked then
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
End if

'Updating the Child Only MA document
If child_only_MA_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Child Only MA.dotx")
	With objDoc
		.FormFields("field_name").Result = CP_name
		.FormFields("field_street_address").Result = street_address
		.FormFields("field_city_state_zip").Result = city_state_zip
		.FormFields("field_case_number").Result = PRISM_case_number
	End With
End if

'Updating the Child Only MA - relative caretaker document
If child_only_MA_relative_caretaker_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Child Only MA - Relative Caretaker.dotx")
	With objDoc
		.FormFields("field_name").Result = CP_name
		.FormFields("field_street_address").Result = street_address
		.FormFields("field_city_state_zip").Result = city_state_zip
		.FormFields("field_case_number").Result = PRISM_case_number
	End With
End if

'Updating the CP and Child MA document
If CP_and_child_MA_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\CP and Child MA.dotx")
	With objDoc
		.FormFields("field_name").Result = CP_name
		.FormFields("field_street_address").Result = street_address
		.FormFields("field_city_state_zip").Result = city_state_zip
		.FormFields("field_case_number").Result = PRISM_case_number
	End With
End if

'Updating the CP paternity request document
If CP_paternity_request_sheet_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\CP Paternity Request Sheet.dotx")
	With objDoc
		.FormFields("field_childs_name").Result = childs_name
		.FormFields("field_CP_name").Result = CP_name
		.FormFields("field_AF_name").Result = NCP_name
		.FormFields("field_case_number").Result = PRISM_case_number
	End With
End if

'Updating the Financial Affidavit OCS document
If financial_affidavit_OCS_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Financial Affidavit OCS.dotx")
	With objDoc
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_all_children").Result = CAPS_kids
		.FormFields("field_CP_name").Result = CP_name
		'Must also add one of these: Name____________________________	Date of Birth_______________ for each kid, and autofill them!!!<<<<<<<<<<<
	End With
End if

'Updating the Normal Paternity Cover Letter to CP document
If paternity_cover_letter_normal_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Paternity Cover letter to CP - Normal.dotx")
	With objDoc
		.FormFields("field_NCP_gender").Result = NCP_gender
		.FormFields("field_NCP_gender_02").Result = NCP_gender
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
		.FormFields("field_phone").Result = worker_phone
	End With
End if

'Opening the Relative Caretaker Paternity Cover Letter to CP document (does not autofill any info)
If paternity_cover_letter_relative_caretaker_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Paternity Cover letter to CP - Relative Caretaker.dotx")
	With objDoc
		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
		.FormFields("field_phone").Result = worker_phone
	End With
End if

'Updating the Minor with GAL Paternity Cover Letter to CP document
If paternity_cover_letter_minor_check = checked then
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Paternity Cover letter to CP - Minor with GAL attachment.dotx")
	With objDoc
		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
		.FormFields("field_phone").Result = worker_phone
	End With
End if

'Opening the Paternity Information Form Memo document (does not autofill any info)
If paternity_information_form_memo_check = checked then set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Paternity Information Form Memo.dotx")

'Opening the Paternity Information Form document (does not autofill any info)
If paternity_information_form_check = checked then 
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Paternity Information Form.dotx")
	With objDoc
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_childs_name").Result = childs_name
		.FormFields("field_childs_name_2").Result = childs_name
	End With
End if

'Opening the Supplemental Paternity Information Form document (does not autofill any info)
If supplemental_paternity_information_form_check = checked then 
	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\Supplemental Paternity Information Form.dotx")
	With objDoc
		.FormFields("field_case_number").Result = PRISM_case_number
		.FormFields("field_childs_name").Result = childs_name
		.FormFields("field_fathers_name").Result = NCP_name
	End With
End if

'If F0018 is indicated on the dialog then it navigates to DORD to send it.
If F0018_check = checked then
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0018", 6, 36
	transmit
End if

'If F0022 is indicated on the dialog then it navigates to DORD to send it.
If F0022_check = checked then
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0022", 6, 36
	transmit
End if

If CAAD_note_check = checked then

	'Going to CAAD, adding a new note
	call navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then script_end_procedure("The script could not navigate to a case note. There may have been a script error. Add case note manually, and report the error to a script writer.")


	'Setting the type
	EMWriteScreen "M2123", 4, 54

	'Setting cursor in write area and writing note details
	EMSetCursor 16, 4
	call write_new_line_in_PRISM_case_note("* Paternity packet sent to CP with the following docs:")
	If child_only_MA_check = checked then call write_new_line_in_PRISM_case_note("    * Child Only MA Choice of Service letter")
	If child_only_MA_relative_caretaker_check = checked then call write_new_line_in_PRISM_case_note("    * Child Only MA - relative caretaker letter")
	If CP_and_child_MA_check = checked then call write_new_line_in_PRISM_case_note("    * CP and Child MA choice of service letter")
	If CP_paternity_request_sheet_check = checked then call write_new_line_in_PRISM_case_note("    * CP Paternity Request sheet")
	If financial_affidavit_OCS_check = checked then call write_new_line_in_PRISM_case_note("    * Financial Affidavit OCS")
	If issues_paternity_to_be_decided_check = checked then call write_new_line_in_PRISM_case_note("    * Issues-Paternity-to be Decided")
	If parenting_time_schedules_check = checked then call write_new_line_in_PRISM_case_note("    * Parenting Time Schedules")
	If paternity_cover_letter_normal_check = checked then call write_new_line_in_PRISM_case_note("    * Normal Paternity Cover Letter to CP")
	If paternity_cover_letter_relative_caretaker_check = checked then call write_new_line_in_PRISM_case_note("    * Relative Caretaker Paternity Cover Letter to CP")
	If paternity_cover_letter_minor_check = checked then call write_new_line_in_PRISM_case_note("    * Minor with GAL Attachment")
	If paternity_information_form_memo_check = checked then call write_new_line_in_PRISM_case_note("    * Paternity Information Form Memo")
	If paternity_information_form_check = checked then call write_new_line_in_PRISM_case_note("    * Paternity Information Form")
	If supplemental_paternity_information_form_check = checked then call write_new_line_in_PRISM_case_note("    * Supplemental Paternity Information Form")
	If F0018_check = checked then call write_new_line_in_PRISM_case_note("    * DORD F0018")
	If F0022_check = checked then call write_new_line_in_PRISM_case_note("    * DORD F0022")
	call write_new_line_in_PRISM_case_note("---")
	call write_new_line_in_PRISM_case_note("* CP to return by " & dateadd("d", date, 5) & ".")
	call write_new_line_in_PRISM_case_note("---")
	call write_new_line_in_PRISM_case_note(worker_signature)

	transmit
End if

If CAWD_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note	
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "*** Pat Docs due from CP", 10, 4
	EMWriteScreen dateadd("d", date, 7), 17, 21
	transmit
End if

script_end_procedure("")

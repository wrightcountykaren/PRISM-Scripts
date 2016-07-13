'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MOD CAAD NOTE: CONTACT CHECKLIST.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denomination = "C"
'END OF STATS BLOCK-------------------------------------------------------------------------------------------------


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

BeginDialog Modification_Case_Note, 0, 0, 371, 440, "MOD CAAD NOTE: CONTACT CHECKLIST"
  Text 15, 10, 50, 10, "Contact Type"
  DropListBox 80, 10, 175, 15, "Select one:"+chr(9)+"T0050 PHONE CALL TO CP"+chr(9)+"T0051 PHONE CALL FR CP"+chr(9)+"T0052 PHONE CALL RET TO CP"+chr(9)+"T0053 PHONE CALL RET FR CP"+chr(9)+"T0101 PHONE CONTACT CP'S ATTORNEY"+chr(9)+"T0093 CONTACT WITH CP SPOUSE"+chr(9)+"M3910 INTERVIEW WITH CP"+chr(9)+"T0055 PHONE CALL TO NCP"+chr(9)+"T0056 PHONE CALL FR NCP"+chr(9)+"T0057 PHONE CALL RET TO NCP"+chr(9)+"T0058 PHONE CALL RET FR NCP"+chr(9)+"T0065 PHONE CALL TO NCP AY"+chr(9)+"T0066 PHONE CALL FR NCP AY"+chr(9)+"T0092 CONTACT WITH NCP SPOUSE"+chr(9)+"M3911 INTERVIEW WITH NCP", Contact_Type_dropdown
  Text 15, 35, 65, 10, "Who Requested:"
  EditBox 80, 30, 130, 15, Who_requested_editbox
  Text 15, 55, 50, 10, "Case Number:"
  EditBox 80, 50, 130, 15, PRISM_case_number
  Text 15, 80, 80, 10, "What is the change?"
  DropListBox 105, 75, 135, 15, "Select one:"+chr(9)+"Income Change"+chr(9)+"Child Emancipates"+chr(9)+"Child Care Change"+chr(9)+"Insurance Change"+chr(9)+"Other ", Change_Options_droplist
  Text 15, 100, 245, 10, "If change is 'income change' or 'other'  please provide more information:"
  EditBox 10, 115, 350, 15, more_info_editbox
  CheckBox 15, 135, 250, 10, "Inform the Client that the support could go up, down or remain the same.", Up_Down_Same_checkbox
  CheckBox 15, 153, 190, 10, "Inform the Client of the online child support calculator", Online_calculator_checkbox
  CheckBox 15, 170, 230, 10, "Inform the Client that once the review has started we cannot stop.", Cannot_stop_checkbox
  Text 15, 195, 170, 10, "Amount of time informed that an Agency Mod takes."
  EditBox 190, 190, 170, 15, Amt_time_editbox
  CheckBox 15, 215, 295, 10, "Inform the Client that the Effective Date is the month following the service date", Effective_date_checkbox
  CheckBox 15, 232, 220, 10, "Inform the Client their option to complete a Pro Se Modification", Pro_se_checkbox
  CheckBox 15, 255, 60, 10, "Verify address", Verify_address_checkbox
  EditBox 85, 250, 275, 15, Address_editbox
  CheckBox 15, 280, 95, 10, "Verify telephone number", Verify_number_checkbox
  EditBox 115, 275, 95, 15, Phone_number_editbox
  CheckBox 15, 302, 105, 10, "Verify employer (NCID/CPID)", Verify_Employer_checkbox
  EditBox 125, 298, 235, 15, Employer_editbox
  CheckBox 15, 325, 150, 10, "Verify e-mail address (use * instead of @)", Verify_email_checkbox
  EditBox 165, 320, 195, 15, Email_editbox
  CheckBox 15, 345, 345, 10, "*Offer the option to fill out financial statement electronically, inform won't start review until received.", Electronic_financial_statements_checkbox
  Text 15, 370, 65, 10, "Other discussions:"
  EditBox 95, 365, 265, 15, Other_editbox
  Text 205, 385, 65, 10, "Worker's Signature"
  EditBox 200, 400, 160, 15, Workers_Signature
  ButtonGroup ButtonPressed
    OkButton 255, 420, 50, 15
    CancelButton 310, 420, 50, 15
EndDialog


'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out
CALL check_for_PRISM(True)

'Searches for the case number.
Call PRISM_case_number_finder(PRISM_case_number)

'Displays dialog for Modification caad note and checks for information

Do
'Shows dialog, validates PRISM mandated fields completed, with transmit
	err_msg = ""
	Dialog Modification_Case_Note
	cancel_confirmation	
	CALL Prism_case_number_validation(prism_case_number, case_number_valid)
	IF Contact_Type_dropdown = "" THEN err_msg = err_msg & vbNEWline & "You must select a contact type!"
	IF Workers_Signature = "" THEN err_msg = err_msg & vbNEWline & "You must sign your CAAD note"
	IF Who_requested_editbox = "" THEN err_msg = err_msg & vbNEWline & "You must enter in who you discussed Modification Options with!"
	IF (Change_Options = "Income Change (ask why)" or Change_Options = "Other") and Change_type = "" THEN err_msg = err_msg & vbNEWline & "Please provide more detail!"
	IF err_msg <> "" THEN MsgBox "***Notice***" & vbNEWline & err_msg &vbNEWline & vbNEWline & "Please resolve for the script"
LOOP UNTIL err_msg = ""	


'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")


'Entering case number
call enter_PRISM_case_number(PRISM_case_number, 20, 8)


PF5					'Did this because you have to add a new note

EMWriteScreen Left(Contact_Type_dropdown, 5), 4, 54  'adds correct caad code 

EMSetCursor 16, 4			'Because the cursor does not default to this location

call write_editbox_in_PRISM_case_note("Who discussed the Modification with", Who_requested_editbox, 4)
call write_editbox_in_PRISM_case_note("What is the Change", Change_Options_droplist, 4)
call write_editbox_in_PRISM_case_note("More info on change", more_info_editbox, 4)
IF Up_Down_Same_checkbox = 1 then call write_new_line_in_PRISM_case_note("* Informed of the online child support calculator.")
IF Online_calculator_checkbox = 1 then call write_new_line_in_PRISM_case_note("* Informed that the support could go up, down or remain the same.")
IF Cannot_stop_checkbox = 1 then call write_new_line_in_PRISM_case_note("* Informed that once a review has started we cannot stop it.")
call write_editbox_in_PRISM_case_note("Total timeframe to complete the modification given", Amt_time_editbox, 4)
IF Effective_date_checkbox = 1 then call write_new_line_in_PRISM_case_note("* Informed that the Effective Date is the month following service date.")
IF Pro_se_checkbox = 1 then call write_new_line_in_PRISM_case_note("* Informed of their option to complete a Pro Se Modification.")
IF Verify_address_checkbox = 1 then call write_editbox_in_PRISM_case_note("Verified and updated Client's address", Address_editbox, 4)
IF Verify_number_checkbox = 1 then call write_editbox_in_PRISM_case_note("Verified and updated Client's phone number", Phone_number_editbox, 4)
IF Verify_employer_checkbox = 1 then call write_editbox_in_PRISM_case_note("Verified and updated Client's employers", Employer_editbox, 4)
IF Verify_email_checkbox = 1 then call write_editbox_in_PRISM_case_note("Verified and updated Client's email", Email_editbox, 4)
IF Electronic_financial_statements_checkbox = 1 then call write_new_line_in_PRISM_case_note("Offered financial statement electronically. Advised would not start review until received back.")
call write_editbox_in_PRISM_case_note("Other discussions", Other_editbox, 4)	
call write_new_line_in_PRISM_case_note(Workers_Signature)

script_end_procedure("")

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - FEE SUSPENSION OVERRIDE.vbs"
start_time = timer
'MANUAL TIME TO COMPLETE THIS SCRIPT IS NEEDED

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


'DIMMING variables
DIM beta_agency, row, col, worker_signature, ButtonPressed, Fee_Suppression_dialog, PRISM_case_number, Fee_date, CAAD_standard_checkbox,CAAD_text_checkbox, CAAD_text 

'THE DIALOG----------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog Fee_Suppression_dialog, 0, 0, 346, 170, "Fee Code Suppression"
  EditBox 45, 25, 75, 15, PRISM_Case_number
  EditBox 65, 45, 50, 15, Fee_date
  CheckBox 15, 90, 320, 10, "Supervisor suppressed cost recovery fee.  Case is NPA due to MNSURE/METS interface issue.  ", CAAD_standard_checkbox
  CheckBox 15, 110, 90, 10, "Enter text for CAAD note.", CAAD_text_checkbox
  EditBox 110, 105, 225, 15, CAAD_text
  EditBox 75, 130, 55, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 235, 150, 50, 15
    CancelButton 290, 150, 50, 15
  Text 5, 10, 280, 10, "This script will manually override the Fee Code Suppression and create a CAAD note."
  Text 15, 30, 30, 10, "Case #:"
  Text 15, 50, 50, 10, "Fee Elig Date: "
  Text 125, 50, 70, 10, "(format 01/01/2001)"
  Text 5, 75, 105, 10, "Select CAAD note option below."
  Text 15, 135, 60, 10, "Worker Signature"
EndDialog


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to Bluezone
EMConnect ""			

'brings me to the CAPS screen
CALL navigate_to_PRISM_screen ("CAST")

'this auto fills prism case number in dialog
EMReadScreen PRISM_case_number, 13, 4, 8 


'adding LOOP to make sure info in dialog box is entered correctly
DO
	err_msg = ""
	dialog Fee_Suppression_dialog
	IF buttonpressed = 0 THEN stopscript
	IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "Please sign your CAAD Note"				'if the signature is blank pop up a message box
	IF Fee_date = "" THEN err_msg = err_msg & vbNewline & "Fee Eligibility end date must be entered."
	IF CAAD_standard_checkbox = 0 AND CAAD_text_checkbox = 0 THEN err_msg = err_msg & vbNewline & "Please select one CAAD note option."
	IF CAAD_standard_checkbox = 1 AND CAAD_text_checkbox = 1 THEN err_msg = err_msg & vbNewline & "Please select only one CAAD note option."
	IF CAAD_text_checkbox = 1 AND CAAD_text = "" THEN err_msg = err_msg & vbNewline & "Please enter the text for your CAAD note."
 	IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
	END IF

LOOP UNTIL err_msg = "" 							                     	

'END LOOP


'Checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)

'fixes date to the correct format xx/xx/xxxx
CALL create_mainframe_friendly_date(FEE_date, 10, 17, "YYYY")

'Goes to CAST screen and PF11 over 							
CALL navigate_to_PRISM_screen("CAST")										
PF11																

'Updates State Fee Cd: to M in order to suppress the 2% fee and adds date
EMWritescreen "M", 9, 17
EMSetCursor 10, 17									        			
EMWritescreen FEE_date, 10, 17
EMWritescreen "M", 3, 29
transmit

'Writes info into CAAD for standard note
IF CAAD_standard_checkbox = 1 THEN 		
	CALL Navigate_to_PRISM_screen ("CAAD")										'navigates to CAADescreen "FREE", 4, 54												'types title of the free caad on the first line of the note	
	PF5
	EMWriteScreen "Free", 4, 54
	EMWriteScreen "*Cost Recovery Fee Override*", 16, 4								'writes this as a title line for the caad note.
	EMSetCursor 17, 4													                    
	CALL write_variable_in_CAAD ("Supervisor suppressed cost recovery fee until "  &  Fee_date &  ".  Case is NPA due to MNSure/METS interface issue.")  
	CALL write_variable_in_CAAD(worker_signature)							  		'adds worker initials from dialog box
	transmit 
	PF3
END IF

IF CAAD_text_checkbox = 1 THEN
	CALL Navigate_to_PRISM_screen ("CAAD")										'navigates to CAADescreen "FREE", 4, 54												'types title of the free caad on the first line of the note	
	PF5
	EMWriteScreen "Free", 4, 54
	EMWriteScreen "*Cost Recovery Fee Override*", 16, 4								'writes this as a title line for the caad note.
	EMSetCursor 17, 4												
	CALL write_variable_in_CAAD(CAAD_text) 
	CALL write_variable_in_CAAD(worker_signature)							  		'adds worker initials from dialog box
	transmit
	PF3
END IF

script_end_procedure("")                                                                     	'stopping the script


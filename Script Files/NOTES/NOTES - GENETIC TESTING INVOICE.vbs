'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - GENETIC TESTING INVOICE.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 60
STATS_denomination = "C"
'END OF STATS BLOCK-------------------------------------------------------------------------------------------


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


'Calling dialog details for the Genetic Testing Invoice---------------------------------------------------------------------

BeginDialog genetic_test_invoice, 0, 0, 306, 115, "Genetic Testing Invoice"
  EditBox 85, 5, 65, 15, prism_case_number
  EditBox 230, 5, 65, 15, invoice_recd_date
  EditBox 85, 25, 90, 15, invoice_from
  EditBox 230, 25, 65, 15, invoice_number
  EditBox 85, 45, 65, 15, dollar_amount
  CheckBox 195, 50, 80, 10, "Invoice is ok to pay", pay_yes_checkbox
  EditBox 40, 65, 245, 15, Edit7
  EditBox 90, 90, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 90, 50, 15
    CancelButton 245, 90, 50, 15
  Text 15, 95, 70, 10, "Sign your CAAD note:"
  Text 15, 30, 65, 10, "Invoice Rec'd From:"
  Text 10, 10, 70, 10, "PRISM Case Number:"
  Text 180, 10, 50, 10, "Invoice Rec'd:"
  Text 195, 30, 35, 10, "Invoice #:"
  Text 15, 70, 25, 10, "Notes:"
  Text 45, 50, 35, 10, "$ Amount:"
EndDialog


'Connecting to Bluezone
EMConnect ""			

'Searches for the case number
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if


'The script will not run unless the CAAD note is signed and there is a valid prism case number
DO
	err_msg = ""
	Dialog genetic_test_invoice
	IF ButtonPressed = 0 THEN StopScript		                                       'Pressing Cancel stops the script
	CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF IsDate(invoice_recd_date) = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid date!"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You sign your CAAD note!"                   
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""


'Checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)


'Goes to CAAD
CALL Navigate_to_PRISM_screen ("CAAD")										'goes to the CAAD screen
PF5																'F5 to add a note
EMWritescreen "A", 3, 29												'put the A on the action line

'Writes info from dialog into CAAD
EMWritescreen "FREE", 4, 54												'types free on caad code: line
EMWritescreen "Genetic Testing Invoice", 16, 4								      'types title of the free caad on the first line of the note
EMSetCursor 17, 4														'puts the cursor on the very next line to be ready to enter the info

CALL write_bullet_and_variable_in_CAAD("Invoice Rec'd", invoice_recd_date)
CALL write_bullet_and_variable_in_CAAD("Invoice Rec'd From", invoice_from)
call write_bullet_and_variable_in_CAAD("invoice #",invoice_number)
call write_bullet_and_variable_in_CAAD("$",dollar_amount)
If pay_yes_checkbox = 1 then call write_variable_in_CAAD("Invoice is OK to pay")
If pay_yes_checkbox = 0 then call write_variable_in_CAAD("Do Not pay invoice")
CALL write_bullet_and_variable_in_CAAD("Notes", notes)
call write_variable_in_CAAD(worker_signature)
transmit
PF3

script_end_procedure("")                                                                     	'stopping the script


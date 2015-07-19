option explicit

DIM beta_agency, row, col

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


'Searches for the case number.
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if


DIM financial_stmt, cover_letter, waiver, calendar, past_support, return_date, signature, PRISM_case_number, mes_caad_dialog, buttonpressed, case_number_valid

'Calling dialog details for the MES CAAD Note---------------------------------------------------------------------
BeginDialog MES_CAAD_dialog, 0, 0, 221, 180, "MES CAAD Note"
  EditBox 70, 5, 90, 15, prism_case_number
  CheckBox 15, 35, 75, 15, "Financial Statement", financial_stmt
  CheckBox 15, 50, 80, 15, "Cover Letter", Cover_letter
  CheckBox 15, 65, 100, 15, "Waiver of Personal Service", Waiver
  CheckBox 15, 80, 95, 15, "Parenting Time Calendar", calendar
  CheckBox 15, 95, 70, 15, "Past Support Form", Past_support
  EditBox 100, 115, 60, 15, return_date
  EditBox 100, 135, 60, 15, signature
  ButtonGroup ButtonPressed
    OkButton 105, 155, 50, 15
    CancelButton 160, 155, 50, 15
  Text 15, 10, 50, 10, "Prism Case #:"
  Text 25, 25, 140, 10, "What documents were sent to the parties?"
  Text 15, 120, 80, 10, "Requested Return Date:"
  Text 45, 140, 50, 10, "Worker initials:"
EndDialog


'Connecting to Bluezone
EMConnect ""			

'checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)

DO																'inserting the loop so that the date and signature are required fields (to start the loop type DO)
	Dialog MES_CAAD_dialog												'open the dialog box itself
	IF ButtonPressed = 0 THEN StopScript				                      'if cancel button is pressed, the script will stop running
	IF signature = "" THEN MSGbox "Please sign your CAAD Note"						'if the signature is blank pop up a message box
	IF return_date = "" THEN MSGbox "Please enter a requested return date"				'if the date field is blank pop up a message box
		IF IsDate(return_date) = False THEN MsgBox "You must enter a valid date"		'makes sure the date field is a valid date
LOOP UNTIL signature <>"" and return_date <>"" and IsDate(return_date) = TRUE				'tells the loop to keep running until the date and signature fields are filled in and the date is valid.  (if you have a Do stmt, you must have a LOOP UNTIL stmt)
		

'go to CAAD
CALL Navigate_to_PRISM_screen ("CAAD")										'goes to the CAAD screen
PF5																'F5 to add a note
EMWritescreen "A", 3, 29												'put the A on the action line

'writes info from dialog into caad
EMWritescreen "FREE", 4, 54												'types free on caad code: line
EMWritescreen "MES Documents sent:", 16, 4									'types title of the free caad on the first line of the note
EMSetCursor 17, 4														'puts the cursor on the very next line to be ready to enter the ate

If financial_stmt = 1 then call write_variable_in_CAAD(" - Financial Statement")			'putting the info that is check from the dialog box into the caad if it is checked
If cover_letter = 1 then call write_variable_in_CAAD(" - Cover Letter")
If waiver = 1 then call write_variable_in_CAAD(" - Waiver of Personal Service")
If calendar = 1 then call write_variable_in_CAAD(" - Parenting Time Calendar")
If past_support = 1 then call write_variable_in_CAAD(" - Past Support Due form")
call write_bullet_and_variable_in_CAAD(" Requested return date", return_date)
call write_variable_in_CAAD(signature)
transmit
PF3

'go to CAWD to make a free worklist 30 days from today
CALL Navigate_to_PRISM_screen ("CAWD")
PF5
EMWritescreen "A", 3, 30

'creates the free worklist to show up in 30 days to check for the docs
EMWritescreen "FREE", 4, 37												'types free on worklist item: line
EMWritescreen "Have MES documents been returned?", 10, 4 						      'types description, have docs been returned   
EMWritescreen "21", 17, 52      			                  					'types 30 in the calendar days field
transmit
PF3

script_end_procedure("")                                                                     	'stopping the script



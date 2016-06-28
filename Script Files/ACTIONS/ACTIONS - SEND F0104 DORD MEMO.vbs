'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - SEND F0104 DORD MEMO.vbs"
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



'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

BeginDialog memo_dialog, 0, 0, 187, 86, "DORD Memo Dialog"
  DropListBox 10, 20, 110, 20, "Select One"+chr(9)+"CPP - Custodial Parent"+chr(9)+"NCP - Noncustodial Parent"+chr(9)+"BOTH - CP and NCP"+chr(9)+"CPE - CP's Employer"+chr(9)+"NCE - NCP's Employer", recipient
  EditBox 10, 60, 90, 14, memo_text
  ButtonGroup ButtonPressed
    PushButton 140, 10, 40, 14, "Preview", preview_button
    OkButton 140, 30, 40, 14
    CancelButton 140, 50, 40, 14
  Text 10, 10, 90, 10, "Select your recipient:"
  Text 10, 40, 90, 20, "Enter the memo text for your F0104 DORD Memo:"
EndDialog



'CUSTOM FUNCTION----------------------------------------------------------------------------------------------

FUNCTION write_value_and_transmit(text, row, col)
	EMWriteScreen text, row, col
	transmit
END FUNCTION

FUNCTION write_text_to_DORD(string_to_write, recipient)
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0104", 6, 36
	EMWriteScreen recipient, 11, 51
	transmit
	
	'This function will add a string to DORD docs.
	IF len(string_to_write) > 1080 THEN 
		MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & _
				"The text below is longer than the script can handle in one DORD document. The script will not add the text to the document." & vbCr & vbCr & _
				string_to_write
		EXIT FUNCTION
	END IF

	dord_rows_of_text = Int(len(string_to_write) / 60) + 1
	
	
	ReDim write_array(dord_rows_of_text)
	'Splitting the text
	string_to_write = split(string_to_write)
	array_position = 1
	FOR EACH word IN string_to_write
		IF len(write_array(array_position)) + len(word) <= 60 THEN 
			write_array(array_position) = write_array(array_position) & word & " "
		ELSE
			array_position = array_position + 1
			write_array(array_position) = write_array(array_position) & word & " "
		END IF
	NEXT
		
	PF14

	'Selecting the "U" label type
	CALL write_value_and_transmit("U", 20, 14)

	'Writing the values
	dord_row = 7
	FOR i = 1 TO dord_rows_of_text
		CALL write_value_and_transmit("S", dord_row, 5)
		
		CALL write_value_and_transmit(write_array(i), 16, 15)
		
		dord_row = dord_row + 1
		IF i = 12 THEN 
			PF8
			dord_row = 7
		END IF
	NEXT
	PF3
	EMWriteScreen "M", 3, 29
	transmit


	
END FUNCTION
'THE SCRIPT----------------------------------------------------------------------------------------------------
 
'Connects to BlueZone
EMConnect ""

'Finds the PRISM case number using a custom function
CALL PRISM_case_number_finder(PRISM_case_number)

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		IF buttonpressed = 0 THEN stopscript
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	LOOP UNTIL case_number_valid = TRUE
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	IF PRISM_check <> "PRISM" THEN MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
LOOP UNTIL PRISM_check = "PRISM"

'Clearing case info from PRISM
CALL navigate_to_PRISM_screen("REGL")
transmit

'Navigating to CAPS
CALL navigate_to_PRISM_screen("CAPS")


'Entering case number and transmitting
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit	
														'Transmitting into it


'Shows dialog, checks to make sure we're still in PRISM (not passworded out)
DO
	error_msg = ""
	Dialog memo_dialog
	IF buttonpressed = 0 THEN stopscript
	
	
	IF buttonpressed = preview_button THEN
		
		IF recipient = "Select One" THEN
			error_msg = error_msg & "Please specify the memo recipient.  "
		END IF
		IF trim(memo_text) = "" THEN
			error_msg = error_msg & "Please enter the memo text. "
		END IF
		IF error_msg <> "" THEN
			msgbox error_msg & "Please resolve to continue."
		ELSE
			msgbox "Recipient: " & recipient & vbCr & vbCr & "*** PREVIEW *** " & vbCr & memo_text
		END IF
	End IF


LOOP UNTIL buttonpressed <> preview_button and error_msg = ""

DO
	error_msg = ""

	IF recipient = "Select One" THEN
		error_msg = error_msg & "Please specify the memo recipient.  "
	END IF
	IF trim(memo_text) = "" THEN
		error_msg = error_msg & "Please enter the memo text. "
	END IF
	IF error_msg <> "" THEN
		msgbox error_msg & "Please resolve to continue."
		Dialog memo_dialog
	END IF
LOOP UNTIL error_msg = ""
	
	
check_for_PRISM(false)

IF recipient = "BOTH - CP and NCP" THEN
	memo_text_for_CP = memo_text
	memo_text_for_NCP = memo_text

	CALL write_text_to_DORD (memo_text_for_CP, "CPP")
	CALL write_text_to_DORD (memo_text_for_NCP, "NCP")
ELSE
	recipient_code = left(recipient, 3)
	CALL write_text_to_DORD (memo_text, recipient_code)
END IF
script_end_procedure("")





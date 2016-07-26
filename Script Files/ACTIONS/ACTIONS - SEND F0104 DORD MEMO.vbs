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

buffer_size = 5 'number of lines to buffer when creating the array.  Due to wrapping, the array may need more lines than initially projected.


'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

BeginDialog memo_dialog, 0, 0, 347, 106, "DORD Memo Dialog"
  DropListBox 10, 20, 110, 20, "Select One"+chr(9)+"CPP - Custodial Parent"+chr(9)+"NCP - Noncustodial Parent"+chr(9)+"BOTH - CP and NCP"+chr(9)+"CPE - CP's Employer"+chr(9)+"NCE - NCP's Employer", recipient_code
  EditBox 10, 60, 270, 14, memo_text
  ButtonGroup ButtonPressed
    PushButton 290, 30, 40, 14, "Preview", preview_button
    PushButton 290, 10, 40, 14, "SpellCheck", spell_button
    OkButton 290, 50, 40, 14
    CancelButton 290, 70, 40, 14
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
	string_to_write_length = len(string_to_write)
	IF string_to_write_length > 1080 THEN 
		excess_string_text = string_to_write_length - 1080
		MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & _
				"The text is longer than the script can handle in one DORD document. Here is your text:" & vbCr & vbCr & _
				Left(string_to_write, 1080) & vbCR & vbCR & " The following text exceeds the capacity of the document:" & _
				vbCR & vbCr & Right(string_to_write, excess_string_text) & vbCR & vbCr & "Please edit your document text."
		EXIT FUNCTION
	END IF

	
	ReDim write_array(18) 'number of lines available to write
	'Splitting the text
	string_to_write = split(string_to_write)
	array_position = 1
	FOR EACH word IN string_to_write
		IF len(write_array(array_position)) + len(word) <= 60 THEN 
			write_array(array_position) = write_array(array_position) & word & " "
		ELSE
			array_position = array_position + 1
			IF array_position > 18 THEN
				MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & _
				"The text is longer than the script can handle in one DORD document.  " _
				& "Please revise your document text."
				EXIT FUNCTION
			END IF
			write_array(array_position) = write_array(array_position) & word & " "
		END IF

	NEXT
		
	PF14

	'Selecting the "U" label type
	CALL write_value_and_transmit("U", 20, 14)

	'Writing the values
	dord_row = 7
	FOR i = 1 TO array_position
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

FUNCTION write_text_to_msgbox(message_text, recipient)
'Preview memo text in a message box or display error message.
	IF recipient = "Select One" THEN
		error_msg = error_msg & "Please specify the memo recipient.  "
	END IF
	IF trim(message_text) = "" THEN
		error_msg = error_msg & "Please enter the memo text. "
	END IF
	IF error_msg <> "" THEN
		msgbox error_msg & "Please resolve to continue."
	ELSE
	
	message_length = len(message_text)
	
	IF message_length > 1080 THEN 
	excess_message_text = message_length - 1080
		MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & _
				"The text is longer than the script can handle in one DORD document. Here is your text:" & vbCr & vbCr & _
				Left(message_text, 1080) & vbCR & vbCR & " The following text exceeds the capacity of the document:" & _
				vbCR & vbCr & Right(message_text, excess_message_text) & vbCR & vbCr & "Please edit your document text."
		EXIT FUNCTION
	END IF

	msg_rows_of_text = Int(message_length / 60) + 1
	
	ReDim write_array(18) 'Number of rows available for writing
	'Splitting the text
	message_text = split(message_text)
	array_position = 1
	FOR EACH word IN message_text
		IF len(write_array(array_position)) + len(word) <= 60 THEN 
			write_array(array_position) = write_array(array_position) & word & " "
		ELSE
			array_position = array_position + 1
			IF array_position > 18 THEN
				MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & _
				"The text is longer than the script can handle in one DORD document.  " _
				& "Please revise your document text."
				EXIT FUNCTION
			END IF
			write_array(array_position) = write_array(array_position) & word & " "
		END IF

	NEXT
	
	msgbox_text =  "Recipient: " & recipient & vbCr & vbCr & "*** PREVIEW *** " & vbCr
	FOR ii = 1 TO array_position
			msgbox_text = msgbox_text & write_array(ii) & vbCr
	NEXT
	msgbox msgbox_text
	END IF
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
	

	IF buttonpressed = spell_button THEN
		
		'Copy memo text to a new Word document, run spell check, and return the spell checked text to the dialog, close the Word doc
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = TRUE
		SET objDoc = objWord.Documents.Add()
		SET objSel = objWord.Selection
		objSel.TypeText memo_text
		objDoc.CheckGrammar
		objSel.WholeStory
		modified_text = objSel.Text 
		memo_text = modified_text
		objDoc.Close(0)
	End IF

	
	IF buttonpressed = preview_button THEN
		message_text = memo_text
		CALL write_text_to_msgbox(message_text, recipient_code)
	End IF


LOOP UNTIL buttonpressed <> preview_button and buttonpressed <> spell_button and error_msg = ""


'Ensuring that all required fields are completed before continuing with export to DORD.
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

'Export information to DORD doc based on recipient selection.
IF recipient_code = "BOTH - CP and NCP" THEN
	memo_text_for_CP = memo_text
	memo_text_for_NCP = memo_text

	CALL write_text_to_DORD (memo_text_for_CP, "CPP")
	CALL write_text_to_DORD (memo_text_for_NCP, "NCP")
ELSE
	recipient = left(recipient_code, 3)
	CALL write_text_to_DORD (memo_text, recipient)
END IF
script_end_procedure("")

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - PAY OR REPORT.vbs"
start_time = timer
'
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


'case number dialog-
BeginDialog case_number_dialog, 0, 0, 176, 85, "Case number dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog

BeginDialog Pay_or_report_dialog, 0, 0, 291, 215, "Pay or Report"
  Text 10, 10, 40, 10, "Order date:"
  EditBox 50, 10, 60, 15, Order_date
  Text 120, 10, 55, 15, "County Attorney"
  DropListBox 175, 10, 110, 15, "Select one..."+chr(9)+"Michael S. Barone"+chr(9)+"Tonya D.F. Berzat"+chr(9)+"Paul C. Clabo"+chr(9)+"Dorrie B. Estebo"+chr(9)+"Kay M. Gavinski"+chr(9)+"Rachel Morrison"+chr(9)+"Brett Schading"+chr(9)+"D. Marie Sieber", CAO_list
  Text 5, 45, 60, 10, "Purge Condition:"
  EditBox 70, 40, 165, 15, purge_condition
  Text 5, 70, 80, 10, "First payment due:"
  EditBox 85, 70, 45, 15, first_payment_due
  Text 140, 70, 45, 10, "Report date:"
  EditBox 180, 70, 40, 15, report_date_one
  Text 5, 85, 75, 10, "Second payment due:"
  EditBox 85, 85, 45, 15, second_payment_due
  Text 140, 85, 40, 15, "Report date:"
  EditBox 180, 85, 40, 15, report_date_two
  Text 5, 100, 75, 15, "Third payment due:"
  EditBox 85, 100, 45, 15, third_payment_due
  Text 140, 100, 40, 15, "Report date:"
  EditBox 180, 100, 40, 15, report_date_three
  Text 5, 115, 75, 15, "Fourth payment due:"
  EditBox 85, 115, 45, 15, fourth_payment_due
  Text 140, 115, 40, 10, "Report date:"
  EditBox 180, 115, 40, 15, report_date_four
  Text 5, 130, 75, 10, "Fifth payment due:"
  EditBox 85, 130, 45, 15, fifth_payment_due
  Text 140, 130, 40, 10, "Report date:"
  EditBox 180, 130, 40, 15, report_date_five
  Text 5, 145, 75, 10, "Sixth payment due:"
  EditBox 85, 145, 45, 15, sixth_payment_due
  Text 140, 145, 40, 10, "Report date:"
  EditBox 180, 145, 40, 15, report_date_six
  Text 125, 175, 65, 10, "Worker signature"
  EditBox 185, 170, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 195, 50, 15
    CancelButton 225, 195, 50, 15
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

Do					
	Do
		Do

			'Shows dialog, validates that PRISM is up and not timed out, with transmit
			Dialog pay_or_report_dialog
			If buttonpressed = 0 then stopscript
			If ISDate(first_payment_due) = False or ISDate(second_payment_due)= False or ISDate(third_payment_due)= False or ISDate(fourth_payment_due)= False or ISDate(fifth_payment_due)= False or ISDate(sixth_payment_due)= False or ISDate(report_date_one)= False or ISDate(report_date_two)= False or ISDate(report_date_three)= False or ISDate(report_date_four)= False or ISDate(report_date_five)= False or ISDate(report_date_six)= False then MsgBOx "Please type in ALL dates"
		Loop until ISDate(first_payment_due) = TRUE and ISDate(second_payment_due)= TRUE and ISDate(third_payment_due)= TRUE and ISDate(fourth_payment_due)= TRUE and ISDate(fifth_payment_due)= TRUE and ISDate(sixth_payment_due)= TRUE and ISDate(report_date_one)= TRUE and ISDate(report_date_two)= TRUE and ISDate(report_date_three)= TRUE and ISDate(report_date_four)= TRUE and ISDate(report_date_five)= TRUE and ISDate(report_date_six)= TRUE

		transmit
		EMReadScreen PRISM_check, 5, 1, 36
		If PRISM_check <> "PRISM" then MsgBox "You appear to have timed out, or are out of PRISM. Navigate to PRISM and try again."
	Loop until PRISM_check = "PRISM"
	'Makes sure worker enters in signature
	If worker_signature = "" then MsgBox "Sign your CAAD note"
Loop until worker_signature <> ""
	


'Going to CAWT screen
call navigate_to_PRISM_screen("CAWT")

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, due today"		'adding a line in the worklist
EMWriteScreen Cdate(first_payment_due), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, due today"		'adding a line in the worklist
EMWriteScreen Cdate(second_payment_due), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, due today"		'adding a line in the worklist
EMWriteScreen Cdate(third_payment_due), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, due today"		'adding a line in the worklist
EMWriteScreen Cdate(fourth_payment_due), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, due today"		'adding a line in the worklist
EMWriteScreen Cdate(fifth_payment_due), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, due today"		'adding a line in the worklist
EMWriteScreen Cdate(sixth_payment_due), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, report date"	'adding a line in the worklist
EMWriteScreen Cdate(report_date_one), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, report date"	'adding a line in the worklist
EMWriteScreen Cdate(report_date_two), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, report date"	'adding a line in the worklist
EMWriteScreen Cdate(report_date_three), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, report date"	'adding a line in the worklist
EMWriteScreen Cdate(report_date_four), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, report date"	'adding a line in the worklist
EMWriteScreen Cdate(report_date_five), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

PF5									'adding a note
EMWriteScreen "FREE", 4,37					'adding a worklist
EMSetCursor 10,04							'setting the cursor in the correct location on PRISM	
EMSendkey "Check for purge payments, report date"	'adding a line in the worklist
EMWriteScreen Cdate(report_date_six), 17,21		'creating the worklists in PRISM
transmit								'adding the worklist to CAWT
PF3									'backing out of worklist

'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")

							'Entering case number
EMWriteScreen case_number, 20, 8


PF5							'Did this because you have to add a new note

EMWriteScreen "FREE", 4, 54 			'adds free note 



EMSetCursor 16, 4					'Because the cursor does not default to this location
call write_new_line_in_PRISM_case_note("Pay or Report Information")
call write_editbox_in_PRISM_case_note("Purge Condition", purge_condition, 6)  
call write_editbox_in_PRISM_case_note("Order Date", Order_date, 6)
call write_editbox_in_PRISM_case_note("County Attorney", CAO_list, 6)
call write_new_line_in_PRISM_case_note("---")	
call write_new_line_in_PRISM_case_note(worker_signature)

script_end_procedure("")
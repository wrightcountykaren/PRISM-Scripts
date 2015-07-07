'Gathering stats-------------------------------------------------------------------------------------
name_of_script = "BULK - CASE TRANSFER.vbs"
start_time = timer

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

BeginDialog worker_numbers_dlg, 0, 0, 231, 165, "Enter Worker Numbers"
  Text 10, 10, 210, 10, "Please enter a list of CSO Worker Numbers to transfer cases to. "
  Text 10, 30, 210, 30, "NOTE: You can enter either the 8-digit Worker ID or the 11-digit code (County, Office Team, Position). The script can decipher between the different numbers."
  Text 10, 70, 210, 20, "The script will give you a list of workers that are not found in PRISM."
  Text 10, 100, 125, 10, "Separate each worker with a comma."
  EditBox 10, 120, 210, 15, worker_list
  ButtonGroup ButtonPressed
    OkButton 125, 145, 50, 15
    CancelButton 175, 145, 50, 15
EndDialog

'===== CUSTOM FUNCTION FOR DIALOGS FOR EACH WORKER
FUNCTION create_case_numbers_dlg(i, worker_array)

	BeginDialog case_numbers_dlg, 0, 0, 276, 195, "Enter Case Numbers"
	Text 10, 15, 60, 10, "Worker Number"
	Text 10, 35, 60, 10, "Worker Name"
	Text 75, 15, 85, 10, worker_array(i, 0)
	Text 75, 35, 85, 10, worker_array(i, 1)
	EditBox 10, 55, 80, 15, worker_array(i, 2)
	EditBox 10, 75, 80, 15, worker_array(i, 3)
	EditBox 10, 95, 80, 15, worker_array(i, 4)
	EditBox 10, 115, 80, 15, worker_array(i, 5)
	EditBox 10, 135, 80, 15, worker_array(i, 6)
	EditBox 100, 55, 80, 15, worker_array(i, 7)
	EditBox 100, 75, 80, 15, worker_array(i, 8)
	EditBox 100, 95, 80, 15, worker_array(i, 9)
	EditBox 100, 115, 80, 15, worker_array(i, 10)
	EditBox 100, 135, 80, 15, worker_array(i, 11)
	EditBox 190, 55, 80, 15, worker_array(i, 12)
	EditBox 190, 75, 80, 15, worker_array(i, 13)
	EditBox 190, 95, 80, 15, worker_array(i, 14)
	EditBox 190, 115, 80, 15, worker_array(i, 15)
	EditBox 190, 135, 80, 15, worker_array(i, 16)
	ButtonGroup ButtonPressed
		OkButton 170, 175, 50, 15
		PushButton 220, 175, 50, 15, "STOP SCRIPT", stop_script_button
	EndDialog
	
	DIALOG case_numbers_dlg
		IF ButtonPressed = stop_script_button THEN stopscript
END FUNCTION 


'===== THE SCRIPT =====
EMConnect ""
CALL check_for_PRISM(False)

DIALOG worker_numbers_dlg
	IF ButtonPressed = stop_script_button THEN stopscript
	IF InStr(worker_list, "UUDDLRLRBA") <> 0 THEN 
		developer_mode = True
		MsgBox "Developer mode enabled."
	END IF
	
worker_list = replace(worker_list, " ", "")
worker_list = split(worker_list, ",")

number_of_workers = UBound(worker_list)
ReDim worker_array(number_of_workers, 16)

i = 0
FOR EACH cso_worker IN worker_list
	IF cso_worker <> "" THEN 
		worker_array(i, 0) = cso_worker
		i = i + 1
	END IF
NEXT

FOR i = 0 TO number_of_workers
	IF len(worker_array(i, 0)) = 8 THEN 
		'If the length of the worker number is 8 then the script goes to LIPO to gather the 11-digit worker position number.
		CALL navigate_to_PRISM_screen("LIPO")
		lipo_row = 6
		DO
			EMReadScreen worker_id, 8, lipo_row, 15
			EMReadScreen end_of_data, 11, lipo_row, 32
			IF end_of_data = "End of Data" THEN 
				worker_array(i, 1) = "WORKER NOT FOUND"
				EXIT DO
			END IF
			IF UCASE(worker_id) = UCASE(worker_array(i, 0)) THEN 
				EMReadScreen worker_array(i, 1), 30, lipo_row, 26
				EMWriteScreen "D", lipo_row, 4
				transmit
				EMReadScreen LIPO_county, 3, 4, 10
				EMReadScreen LIPO_office, 3, 5, 10
				EMReadScreen LIPO_team, 3, 6, 10
				EMReadScreen LIPO_position, 2, 7, 12
				worker_array(i, 0) = LIPO_county & LIPO_office & LIPO_team & LIPO_position
				CALL find_variable("Name: ", worker_array(i, 1), 30)
				CALL create_case_numbers_dlg(i, worker_array)
				EXIT DO
			ELSE
				lipo_row = lipo_row + 1
				IF lipo_row = 19 THEN 
					PF8
					lipo_row = 6
				END IF
			END IF
		LOOP		
	ELSEIF len(worker_array(i, 0)) = 11 THEN
		CALL navigate_to_PRISM_screen("CALI")
		EMSetCursor 20, 18
		EMSendKey worker_array(i, 0)
		transmit
		
		EMReadScreen error_message_on_bottom_of_screen, 20, 24, 2
		error_message_on_bottom_of_screen = trim(error_message_on_bottom_of_screen)
		IF error_message_on_bottom_of_screen = "" THEN 
			CALL find_variable("Name: ", worker_array(i, 1), 30)
			CALL create_case_numbers_dlg(i, worker_array)
		ELSEIF error_message_on_bottom_of_screen <> "" THEN 
			worker_array(i, 1) = "WORKER NOT FOUND"
		END IF
	ELSE
		worker_array(i, 1) = "WORKER NOT FOUND"
	END IF
NEXT		

'Navigating to CAAS to let the case transferring begin!!
CALL navigate_to_PRISM_screen("CAAS")

err_workers = ""
FOR i = 0 TO number_of_workers
	IF worker_array(i, 1) = "WORKER NOT FOUND" THEN 
		err_workers = err_workers & vbCr & "     " & worker_array(i, 0) 
	ELSEIF worker_array(i, 1) <> "WORKER NOT FOUND" THEN 
		FOR j = 2 TO 16
			IF worker_array(i, j) <> "" THEN 
				CAAS_county = left(worker_array(i, 0), 3)
				CAAS_office = right(left(worker_array(i, 0), 6), 3)
				CAAS_team = left(right(worker_array(i, 0), 5), 3)
				CAAS_position = right(worker_array(i, 0), 2)
								
				EMWriteScreen "M", 3, 29
				EMWriteScreen left(worker_array(i, j), 10), 4, 8
				EMWriteScreen right(worker_array(i, j), 2), 4, 19
				EMWriteScreen CAAS_county, 9, 20
				EMWriteScreen CAAS_office, 10, 20
				EMWriteScreen CAAS_team, 11, 20
				EMWriteScreen CAAS_position, 12, 20
				
				IF developer_mode = True THEN 
					MsgBox "*** Developer Mode Enabled ***" & vbCr & vbCr & _
						"Transferring Case " & worker_array(i, j) & " to " & worker_array(i, 1)				
				ELSE
					transmit
				END IF
			END IF
		NEXT
	END IF
NEXT

'Displaying the list of workers that were skipped because they could not be found.
IF err_workers <> "" THEN MsgBox ("*** NOTICE!!! ***" & vbCr & vbCr & "The script could not transfer cases to the following worker ID/code(s): " & vbCr & err_workers & vbCr & vbCr & "The script has determined that ID/code is not a valid ID/code assigned to a worker. You may need to reconsider the worker ID/code you selected and try again." & vbCr & vbCr & "If the script erred in its determination of valid worker ID/codes, please report this to your scripts administrator." & vbCr & vbCr & "Thank you.")

script_end_procedure("Success!!")

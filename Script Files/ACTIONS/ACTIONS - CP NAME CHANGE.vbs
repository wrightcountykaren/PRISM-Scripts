'Gathering stats
name_of_script = "Action - CP NAME CHANGE.vbs"
start_time = timer
STATS_Counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'End of stats block

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

'the script---------------------------------------------------------------------------------------------------------------
BeginDialog Name_change_dialog, 0, 0, 191, 110, "CP Name Change"
  EditBox 80, 5, 100, 15, Prism_case_number
  EditBox 80, 25, 100, 15, New_name
  EditBox 80, 45, 100, 15, reason_change
  EditBox 80, 65, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 90, 50, 15
    CancelButton 140, 90, 50, 15
  Text 5, 50, 65, 15, "Reason for change:"
  Text 5, 30, 70, 15, "CP New Last Name:"
  Text 5, 70, 65, 15, "Worker Signature:"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog

' connects to Bluezone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out
CALL check_for_PRISM(True)

'Grabs the case number
call PRISM_case_number_finder(Prism_case_number)
	DO
		err_msg = ""
		Dialog Name_change_dialog
		IF ButtonPressed = 0 THEN StopScript		                                       'Pressing Cancel stops the script
		IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You must sign your CAAD note!" 'If worker sig is blank, message box pops saying you must sign caad note
		CALL PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		IF case_number_valid = False THEN MsgBox "Your case number is not valid. Please make sure it is in the following format: XXXXXXXXXX-XX"
		If err_msg <> "" THEN msgbox "***NOTICE***" & vbcr & err_msg & vbNewline & vbNewline & "Please resolve for this script to continue."
	LOOP UNTIL err_msg = ""



'Navigates to CAST
navigate_to_PRISM_screen("CAST")


'Calls the dialog
Dialog Name_change_dialog

'if cancel button is pressed script is canceled
If buttonpressed = 0 then stopscript

'Navigates to CPDE
call navigate_to_prism_screen ("CPDE")

'hits transmit
transmit

'Enters "M" to modify
EMwritescreen "M", 3, 29

'Clears last name
EMWritescreen "__________________", 8, 8

'Hits transmit
transmit

'Hits tranmit
transmit

'Enters "M" to modify
EMwritescreen "M", 3, 29

'Enters new last name from dialog
EMwritescreen new_name, 8,8

'hits transmit
transmit

'hits transmit
transmit

'Navigates to CAAD
call navigate_to_prism_screen("CAAD")

'Enters "M" to modify
EMwritescreen "M", 8,5

'hits transmit
transmit

emsetcursor 16,4

'Enters info for CAAD note
call write_bullet_and_variable_in_caad("Updated CP Name to", New_name)


'enters info on CAAD note
call write_bullet_and_variable_in_caad("Reason for change", reason_change)

'enters CSO signature
call write_variable_in_caad(worker_signature)

'hits transmit
transmit


call script_end_procedure ("")

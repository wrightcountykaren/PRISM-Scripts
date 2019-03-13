'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "unreimbursed-uninsured-sending-docs.vbs"
start_time = timer

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS---------------------------------------------------------------------------
BeginDialog UnUn_Dialog, 0, 0, 291, 145, "Unreimbursed Uninsured Docs"
  EditBox 60, 45, 90, 15, PRISM_case_number
  DropListBox 190, 65, 60, 45, "Select One..."+chr(9)+"CPP"+chr(9)+"NCP", person_droplistbox
  EditBox 175, 100, 25, 15, Percent
  ButtonGroup ButtonPressed
    OkButton 180, 125, 50, 15
    CancelButton 235, 125, 50, 15
  Text 25, 10, 240, 15, "This script will gernerate DORD DOCS F0944, F0659, and F0945 for collection of Unreimbursed and Uninsured Medical and Dental Expenses."
  Text 5, 50, 50, 10, "Case Number"
  Text 5, 70, 175, 10, "Select who requested Unreimbursed/Uninsured forms"
  Text 5, 105, 165, 10, "Enter the PERCENT owed by non requesting party:"
EndDialog


'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

'brings me to the CAPS screend
CALL navigate_to_PRISM_screen ("CAPS")

'check for prism (password out)before continuing
CALL check_for_PRISM(true)

'this auto fills prism case number in dialog
CALL PRISM_case_number_finder(PRISM_case_number)

'THE LOOP--------------------------------------
'adding a loop
Do
	err_msg = ""
	Dialog UnUn_Dialog 'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed"
		IF Percent = "" THEN err_msg = err_msg & vbNewline & "Percent of Unreimbursed Uninsured Expense must be completed."
		IF person_droplistbox = "Select One..." THEN err_msg = err_msg & vbNewline & "Select who requested the documents."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

LOOP UNTIL err_msg = ""

'END LOOP--------------------------------------


'creates DORD doc for NCP from droplist 
IF person_droplistbox = "NCP" THEN
	CALL navigate_to_PRISM_screen ("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0945", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit

	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0944", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit

	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0659", 6, 36
	EMWriteScreen "ncp", 11, 51
	transmit

	'shift f2, to get to user labels
	PF14
	EMWriteScreen "u", 20,14
	transmit
	EMSetCursor 7, 5
	EMWriteScreen "S", 7, 5

	transmit
	EMWriteScreen Percent, 16, 15
	transmit
	PF3
	EMWriteScreen "M", 3, 29
	transmit

'''dialog used because we need to select legal heading
	BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  	  ButtonGroup ButtonPressed
    	    OkButton 60, 75, 50, 15
    	    CancelButton 115, 75, 50, 15
  	  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  	  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  	  Text 5, 40, 55, 10, "2. Press ENTER"
  	  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
	EndDialog

	Dialog LH_dialog  'name of dialog
  	IF buttonpressed = 0 then stopscript		'Cancel

	CALL write_value_and_transmit ("B", 3, 29)
END IF

'creates DORD doc for CP from droplist 
IF person_droplistbox = "CPP" THEN
	CALL navigate_to_PRISM_screen ("DORD")
	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0945", 6, 36
	EMWriteScreen "cpp", 11, 51
	transmit

	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0944", 6, 36
	EMWriteScreen "cpp", 11, 51
	transmit

	EMWriteScreen "C", 3, 29
	transmit

	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0659", 6, 36
	EMWriteScreen "cpp", 11, 51
	transmit

	'shift f2, to get to user labels
	PF14
	EMWriteScreen "u", 20,14
	transmit
	EMSetCursor 7, 5
	EMWriteScreen "S", 7, 5

	transmit
	EMWriteScreen Percent, 16, 15
	transmit
	PF3
	EMWriteScreen "M", 3, 29
	transmit

'''dialog used because we need to select legal heading
	BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  	  ButtonGroup ButtonPressed
    	    OkButton 60, 75, 50, 15
    	    CancelButton 115, 75, 50, 15
  	  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  	  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  	  Text 5, 40, 55, 10, "2. Press ENTER"
  	  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
	EndDialog

	Dialog LH_dialog  'name of dialog
  	IF buttonpressed = 0 then stopscript		'Cancel

CALL write_value_and_transmit ("B", 3, 29)
END IF

script_end_procedure("")

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
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

'Loading all scripts
If run_locally <> true then
	CALL run_from_GitHub("https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/ALL%20SCRIPTS.vbs")
Else
	CALL run_from_GitHub("C:\PRISM-Scripts\ALL SCRIPTS.vbs")
End if

DIM ButtonPressed, button_placeholder
DIM SIR_instructions_button
DIM Dialog1

Function declare_main_menu(menu_type, script_array)
	BeginDialog Dialog1, 0, 0, 516, 440, menu_type & " Scripts"
	  ButtonGroup ButtonPressed
	 	'This starts here, but it shouldn't end here :)
		vert_button_position = 30
		button_placeholder = 100
		FOR current_script = 0 to ubound(script_array)
			IF InStr(script_array(current_script).script_type, menu_type) <> 0 THEN
								'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
				'FUNCTION		HORIZ. ITEM POSITION								VERT. ITEM POSITION		ITEM WIDTH									ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
				PushButton 		5, 													vert_button_position, 	script_array(current_script).button_size, 	10, 			script_array(current_script).script_name, 			button_placeholder
				Text 			script_array(current_script).button_size + 10, 		vert_button_position, 	500, 										10, 			"--- " & script_array(current_script).description
				'----------
				vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
				'----------
				script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
			END IF
			button_placeholder = button_placeholder + 1
		NEXT
		PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button
		CancelButton 460, 420, 50, 15
	EndDialog
End function

DO
	CALL declare_main_menu("ACTIONS", cs_scripts_array)
	Dialog
	IF ButtonPressed = 0 THEN script_end_procedure("")
	IF ButtonPressed = SIR_instructions_button THEN CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/PRISMscripts/Shared%20Documents/Forms/All%20ACTIONS%20Scripts.aspx")
LOOP UNTIL ButtonPressed <> SIR_instructions_button

'Determining the script selected from the value of ButtonPressed
'Since we start at 100 and then go up, we will simply subtract 100 when determining the position in the array
script_picked = ButtonPressed - 100

'Running the selected script
CALL run_from_GitHub(script_repository & cs_scripts_array(script_picked).script_type & "/" & cs_scripts_array(script_picked).file_name)

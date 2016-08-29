'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - QUICK CAAD.vbs"
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

BeginDialog quick_CAAD_dialog, 0, 0, 196, 105, "Quick CAAD dialog"
  ButtonGroup ButtonPressed
    OkButton 85, 85, 50, 15
    CancelButton 140, 85, 50, 15
    PushButton 5, 5, 30, 10, "M3909", CAAD_M3909_button
    PushButton 5, 90, 75, 10, "search CAAD codes...", search_CAAD_codes_button
  Text 40, 5, 150, 10, "DATE OF HEARING"
EndDialog

BeginDialog quick_CAAD_search_dialog, 0, 0, 356, 120, "Quick CAAD search dialog"
  ButtonGroup ButtonPressed
    OkButton 245, 100, 50, 15
    CancelButton 300, 100, 50, 15
  EditBox 110, 15, 55, 15, CAAD_code_to_search
  GroupBox 5, 5, 165, 30, "Option A"
  Text 10, 20, 95, 10, "Enter a CAAD code to search:"
  EditBox 110, 50, 110, 15, CAAD_description_to_search
  GroupBox 5, 40, 220, 30, "Option B"
  Text 10, 55, 100, 10, "Enter a description to search:"
  GroupBox 230, 5, 120, 90, "Instructions"
  Text 235, 15, 110, 75, "Enter a five-digit CAAD code (ex: M9901) into Option A, or (alternatively) enter a description to search for in Option B. If a description is searched, the script will return any matches one by one, and allow you to indicate whether-or-not it is the correct CAAD code."
EndDialog



EMConnect ""

Dialog quick_CAAD_dialog
If ButtonPressed = cancel then StopScript

If ButtonPressed = search_CAAD_codes_button then
	Dialog quick_CAAD_search_dialog
	If ButtonPressed = cancel then StopScript

	If CAAD_code_to_search <> "" then
		EMSetCursor 4, 54		'Where the code is entered, we need to set the cursor there to read the help details
		PF1						'Loads help
		EMWriteScreen CAAD_code_to_search, 20, 28
		transmit
		EMReadScreen CAAD_code_check, 5, 13, 18
		If CAAD_code_check = CAAD_code_to_search then
			MsgBox "found!"
		Else
			MsgBox "Your CAAD code was not found."
		End if
	End if
End if

'If ButtonPressed = CAAD_M3909_button then EMWriteScreen "M3909", 4, 54

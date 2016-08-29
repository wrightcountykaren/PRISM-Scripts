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

'This is a large do...loop which will constantly re-read the text file in case changes were made. If the script is just being used (and not modified), it will only iterate once through this.
Do

	'Before loading dialogs, needs to scan the My Documents folder for a file called favoriteCAADnotes.txt. If this file is found, details about CAAD notes will be pre-loaded. Otherwise, it won't be.
	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now it determines the favorite CAAD notes
	With (CreateObject("Scripting.FileSystemObject"))																				'Creating an FSO
		If .FileExists(user_myDocs_folder & "favoriteCAADnotes.txt") Then															'If the favoriteCAADnotes.txt file exists...
			Set get_favorite_CAAD_notes = CreateObject("Scripting.FileSystemObject")												'Create another FSO
			Set favorite_CAAD_notes_command = get_favorite_CAAD_notes.OpenTextFile(user_myDocs_folder & "favoriteCAADnotes.txt")	'Open the text file
			favorite_CAAD_notes_raw = favorite_CAAD_notes_command.ReadAll															'Read the text file <<<<<CAN THIS READ ONE LINE
			IF favorite_CAAD_notes_raw <> "" THEN favorite_CAAD_notes_array = split(favorite_CAAD_notes_raw, vbNewLine)				'Split by new lines
			favorite_CAAD_notes_command.Close																						'Closes the file
		END IF
	END WITH



	'THIS IS A DYNAMIC DIALOG! DON'T OPEN ME IN DIALOG EDITOR!!!!
	BeginDialog quick_CAAD_dialog, 0, 0, 200, 105, "Quick CAAD dialog"
	  ButtonGroup ButtonPressed
	    OkButton 90, 85, 50, 15
	    CancelButton 145, 85, 50, 15
		dialog_row = 5																										'Starting here so that the contents display in a pretty manner
		'Iterate through each item in the array determined above, then display them
		For i = 0 to ubound(favorite_CAAD_notes_array)																		'i is a counter in this case
			number_to_pass_to_the_button = 1000 + i																			'Add 1000 to the counter to get a ButtonPressed value we can use
			PushButton 5, dialog_row, 30, 10, left(favorite_CAAD_notes_array(i), 5), number_to_pass_to_the_button			'Show the 5-most characters of the CAAD code (each one is five characters) as a button
			Text 40, dialog_row, 150, 10, right(favorite_CAAD_notes_array(i), len(favorite_CAAD_notes_array(i)) - 7)		'Show the rest as a description
			dialog_row = dialog_row + 15																					'Go up 15 pixels
		Next
		PushButton 5, 90, 80, 10, "search CAAD codes...", search_CAAD_codes_button											'Provides a search feature
	EndDialog

	BeginDialog quick_CAAD_search_dialog, 0, 0, 356, 120, "Quick CAAD search dialog"
	  ButtonGroup ButtonPressed
	    OkButton 245, 100, 50, 15
	    CancelButton 300, 100, 50, 15
	  EditBox 115, 15, 55, 15, CAAD_code_to_search
	  EditBox 115, 50, 105, 15, CAAD_description_to_search
	  GroupBox 5, 5, 170, 30, "Option A"
	  Text 10, 20, 100, 10, "Enter a CAAD code to search:"
	  GroupBox 5, 40, 220, 30, "Option B"
	  Text 10, 55, 100, 10, "Enter a description to search:"
	  GroupBox 230, 5, 120, 90, "Instructions"
	  Text 235, 15, 110, 75, "Enter a five-digit CAAD code (ex: M9901) into Option A, or (alternatively) enter a description to search for in Option B. If a description is searched, the script will return any matches one by one, and allow you to indicate whether-or-not it is the correct CAAD code."
	EndDialog

	'Connects to PRISM
	EMConnect ""

	'Displays the dialog, cancels if cancel is pressed
	Dialog quick_CAAD_dialog
	If ButtonPressed = cancel then StopScript

	'Iterates through the array of favorite CAAD notes, and compares with the ButtonPressed value... If the ButtonPressed = i + 1000 (see above), it will have found the right CAAD code!
	For i = 0 to ubound(favorite_CAAD_notes_array)
		If ButtonPressed = i + 1000 then CAAD_code_to_write = left(favorite_CAAD_notes_array(i), 5)	'If the ButtonPressed = i + 1000, then we found the right CAAD note, and can get the CAAD_code_to_write from the left-most 5 characters
	Next

	'If the user wants to search, we need to guide them toward that. Here's how...
	If ButtonPressed = search_CAAD_codes_button then

		'This do...loop will run until ButtonPressed is cancel, or until the user writes info in either CAAD_code_to_search or CAAD_description_to_search.
		Do
			Dialog quick_CAAD_search_dialog																												'Show the dialog
			If ButtonPressed = cancel then exit do																										'Exit do (so we can get out to the main dialog)
			If CAAD_code_to_search = "" and CAAD_description_to_search = "" then MsgBox "You must enter either a CAAD code or description to search."	'Error message

			'If the CAAD_code_to_search is entered, this will process that first. Otherwise it will go through the description to search.
			If CAAD_code_to_search <> "" then
				EMSetCursor 4, 54																														'Where the code is entered, we need to set the cursor there to read the help details
				PF1																																		'Loads help
				EMWriteScreen CAAD_code_to_search, 20, 28																								'Write the search string at 20, 28 on the screen
				transmit																																'Transmits
				EMReadScreen CAAD_code_check, 5, 13, 18																									'Checks the first response
				If CAAD_code_check = CAAD_code_to_search then																							'If the first response is correct...
					add_to_quick_CAAD = MsgBox("The code was found! Would you like to add this to the Quick CAAD button?", vbYesNo + vbQuestion)		'...ask the user if they want to update...
					If add_to_quick_CAAD = vbNo then exit do																							'...If they don't, take them back to the main dialog...
					If add_to_quick_CAAD = vbYes then																									'...If they do, time to add the file!
						'Needs to open up file, manually modify it to include this function, then close the file and re-run the script
					End if
				Else
					MsgBox "Your CAAD code was not found."
					PF3
				End if
			End if


		Loop until ButtonPressed = cancel or (CAAD_code_to_search <> "" or CAAD_description_to_search <> "")

		ButtonPressed = 0	'Needs to blank out so it goes back to the main dialog

	End if

Loop until ButtonPressed <> cancel

'Now it goes to the selected note
navigate_to_PRISM_screen("CAAD")			'Navigates to CAAD
PF5											'Adds a new note
EMWriteScreen CAAD_code_to_write, 4, 54		'Writes the code

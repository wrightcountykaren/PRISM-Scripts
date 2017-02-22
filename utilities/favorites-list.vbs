'TODO: create hotkey file (line 100)
'TODO: make the hotkey dialog do something (line 100)
'TODO: make sure file names display clearly in the display dialog (line 740)

'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "favorites-list.vbs"
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
call changelog_update("02/22/2017", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'script_end_procedure("Favorites button is coming soon! - Veronica Cary (01/31/2017)")

'This function simply displays a list of hotkeys, and the user can insert screens-to-navigate-to within
function edit_hotkeys
	'Instructional MsgBox
	MsgBox  "This section will add PRISM screens to hotkey combinations!" & vbNewLine & vbNewLine & _
			"To use it, simply insert the four-character PRISM screen you'd like to navigate to when pressing the specific key combination." & vbNewLine & vbNewLine & _
			"So, for example, to navigate to CAAD every time Ctrl-F1 is pressed, simply type ""CAAD"" in the editbox." & vbNewLine & vbNewLine & _
			"When you are finished, the script will add a hotkeys file to your My Documents folder, which will store your choices."

	'A dialog
	BeginDialog hotkey_selection_dialog, 0, 0, 116, 285, "Hotkey Selection Dialog"
	  Text 15, 10, 30, 10, "Hotkey:"
	  Text 55, 5, 55, 20, "PRISM screen to navigate to:"
	  Text 15, 30, 25, 10, "Ctrl-F1:"
	  Text 15, 50, 25, 10, "Ctrl-F2:"
	  Text 15, 70, 25, 10, "Ctrl-F3:"
	  Text 15, 90, 25, 10, "Ctrl-F4:"
	  Text 15, 110, 25, 10, "Ctrl-F5:"
	  Text 15, 130, 25, 10, "Ctrl-F6:"
	  Text 15, 150, 25, 10, "Ctrl-F7:"
	  Text 15, 170, 25, 10, "Ctrl-F8:"
	  Text 15, 190, 25, 10, "Ctrl-F9:"
	  Text 10, 210, 30, 10, "Ctrl-F10:"
	  Text 10, 230, 30, 10, "Ctrl-F11:"
	  Text 10, 250, 30, 10, "Ctrl-F12:"
	  EditBox 55, 25, 55, 15, ctrl_f1_hotkey_choice
	  EditBox 55, 45, 55, 15, ctrl_f2_hotkey_choice
	  EditBox 55, 65, 55, 15, ctrl_f3_hotkey_choice
	  EditBox 55, 85, 55, 15, ctrl_f4_hotkey_choice
	  EditBox 55, 105, 55, 15, ctrl_f5_hotkey_choice
	  EditBox 55, 125, 55, 15, ctrl_f6_hotkey_choice
	  EditBox 55, 145, 55, 15, ctrl_f7_hotkey_choice
	  EditBox 55, 165, 55, 15, ctrl_f8_hotkey_choice
	  EditBox 55, 185, 55, 15, ctrl_f9_hotkey_choice
	  EditBox 55, 205, 55, 15, ctrl_f10_hotkey_choice
	  EditBox 55, 225, 55, 15, ctrl_f11_hotkey_choice
	  EditBox 55, 245, 55, 15, ctrl_f12_hotkey_choice
	  ButtonGroup ButtonPressed
	    OkButton 5, 265, 50, 15
	    CancelButton 60, 265, 50, 15
	EndDialog

	'Show the dialog
	Dialog hotkey_selection_dialog
	If ButtonPressed = cancel then StopScript

	'Edit the hotkey file

	'Somehow program the redirects to look at that file and do the magic

end function

'====================================================================================
'====================================================================================
'This VERY VERY long function contains all of the logic behind editing the favorites.
'====================================================================================
'====================================================================================
function edit_favorites

	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>> SECTION 1 <<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>> The gobbins that happen before the user sees anything. <<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


	'Looks up the script details online (or locally if you're a scriptwriter)

	If run_locally <> true then

		'Creating the object to the URL a la text file
		SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")

		'Building an array of all scripts
		'Opening the URL for the given main menu
		get_all_scripts.open "GET", all_scripts_repo, FALSE
		get_all_scripts.send
		IF get_all_scripts.Status = 200 THEN
			Set filescriptobject = CreateObject("Scripting.FileSystemObject")
			Execute get_all_scripts.responseText
		ELSE
			'If the script cannot open the URL provided...
			MsgBox 	"Something went wrong with the URL: " & all_scripts_repo
			stopscript
		END IF
	ELSE
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(all_scripts_repo)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF

	'Warning/instruction box
	MsgBox  "This section will display a dialog with various scripts on it. Any script you check will be added to your favorites menu. Scripts you un-check will be removed. Once you are done making your selection hit ""OK"" and your menu will be updated. " & vbNewLine & vbNewLine &_
			"Note: you will be unable to edit the list of NEW Scripts and Recommended Scripts."

	'An array containing details about the list of scripts, including how they are displayed and stored in the favorites tag
	'0 => The script name
	'1 => The checked/unchecked status (based on the dialog list)
	'2 => The script category, and a "/" so that it's presented in a URL
	'3 => The proper script file name
	'4 => The hotkey the user has associated with the script

	REDIM scripts_edit_favs_array(ubound(cs_scripts_array), 4)

	'determining the number of each kind of script...by category
	number_of_scripts = 0
	actions_scripts = 0
	bulk_scripts = 0
	calc_scripts = 0
	notes_scripts = 0
	utilities_scripts = 0
	FOR i = 0 TO ubound(cs_scripts_array)
		number_of_scripts = i
		IF cs_scripts_array(i).category = "actions" THEN
			actions_scripts = actions_scripts + 1
		ELSEIF cs_scripts_array(i).category = "bulk" THEN
			bulk_scripts = bulk_scripts + 1
		ELSEIF cs_scripts_array(i).category = "calculators" THEN
			calc_scripts = calc_scripts + 1
		ELSEIF cs_scripts_array(i).category = "notes" THEN
			notes_scripts = notes_scripts + 1
		ELSEIF cs_scripts_array(i).category = "utilities" THEN
	        utilities_scripts = utilities_scripts + 1
	    End if
	NEXT


	'>>> If the user has already selected their favorites, the script will open that file and
	'>>> and read it, storing the contents in the variable name ''user_scripts_array''
	SET oTxtFile = (CreateObject("Scripting.FileSystemObject"))
	With oTxtFile
		If .FileExists(favorites_text_file_location) Then
			Set fav_scripts = CreateObject("Scripting.FileSystemObject")
			Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
			fav_scripts_array = fav_scripts_command.ReadAll
			IF fav_scripts_array <> "" THEN user_scripts_array = fav_scripts_array
			fav_scripts_command.Close
		END IF
	END WITH

	'>>> Determining the width of the dialog from the number of scripts that are available...
	'the dialog starts with a width of 400
	dia_width = 400

	'VKC - removed old functionality to determine dynamically the width. This will need to be redetermined based on the number of scripts, but I am holding off on this until I know all of the content I'll jam in here. -11/29/2016

	'>>> Building the dialog
	BeginDialog build_new_favorites_dialog, 0, 0, dia_width, 440, "Select your favorites"
		ButtonGroup ButtonPressed
			OkButton 5, 5, 50, 15
			CancelButton 55, 5, 50, 15
			PushButton 165, 5, 70, 15, "Reset Favorites", reset_favorites_button
		'>>> Creating the display of all scripts for selection (in checkbox form)
		script_position = 0		' <<< This value is tied to the number_of_scripts variable


		col = 10
		row = 30
		Text col, row, 175, 10, "---------- ACTIONS SCRIPTS ----------"
		row = row + 10

		FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "actions" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 430 THEN
			row = 30
			col = col + 195
		END IF
		Text col, row, 175, 10, "---------- BULK SCRIPTS ----------"
		row = row + 10

		'BULK script laying out
		FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "bulk" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 430 THEN
			row = 30
			col = col + 195
		END IF
		Text col, row, 175, 10, "---------- CALCULATOR SCRIPTS ----------"
		row = row + 10

		'CALCULATOR script laying out
		FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "calculators" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 430 THEN
			row = 30
			col = col + 195
		END IF
		Text col, row, 175, 10, "---------- NOTES SCRIPTS ----------"
		row = row + 10

		'NOTES script laying out
		FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "notes" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 430 THEN
			row = 30
			col = col + 195
		END IF
		Text col, row, 175, 10, "---------- UTILITIES SCRIPTS ----------"
		row = row + 10

		'UTILITIES script laying out
	    FOR i = 0 to ubound(cs_scripts_array)
			IF cs_scripts_array(i).category = "utilities" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 430 THEN
					row = 30
					col = col + 195
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(user_scripts_array, "-", " ")), UCASE(replace(cs_scripts_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, 1) = checked
				ELSE
					scripts_edit_favs_array(script_position, 1) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, 0) = cs_scripts_array(i).script_name
				scripts_edit_favs_array(script_position, 3) = cs_scripts_array(i).file_name
				scripts_edit_favs_array(script_position, 2) = cs_scripts_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 185, 10, scripts_edit_favs_array(script_position, 0), scripts_edit_favs_array(script_position, 1)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT
	EndDialog

	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>> SECTION 2 <<<<<<<<<<<<<<<<<<<<<
	'>>> The gobbins that the user sees and makes do. <<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<

	DO
		DO
			'>>> Running the dialog
			Dialog build_new_favorites_dialog
			'>>> Cancel confirmation
			IF ButtonPressed = 0 THEN
				confirm_cancel = MsgBox("Are you sure you want to cancel? Press YES to cancel the script. Press NO to return to the script.", vbYesNo)
				IF confirm_cancel = vbYes THEN script_end_procedure("~PT: Script cancelled.")
			END IF
			'>>> If the user selects to reset their favorites selections, the script
			'>>> will go through the multi-dimensional array and reset all the values
			'>>> for position 1, thereby clearing the favorites from the display.
			IF ButtonPressed = reset_favorites_button THEN
				FOR i = 0 to number_of_scripts
					scripts_edit_favs_array(i, 1) = unchecked
				NEXT
			END IF
		'>>> The exit condition for the first do/loop is the user pressing 'OK'
		LOOP UNTIL ButtonPressed <> 0 AND ButtonPressed <> reset_favorites_button
		'>>> Validating that the user does not select more than a prescribed number of scripts.
		'>>> Exceeding the limit will cause an exception access violation for the Favorites script when it runs.
		'>>> Currently, that value is 30. That is lower than previous because of the larger number of new scripts. (-Robert, 04/20/2016)
		double_check_array = ""
		FOR i = 0 to number_of_scripts
			IF scripts_edit_favs_array(i, 1) = checked THEN double_check_array = double_check_array & scripts_edit_favs_array(i, 0) & "~"
		NEXT
		double_check_array = split(double_check_array, "~")
		IF ubound(double_check_array) > 29 THEN MsgBox "Your favorites menu is too large. Please limit the number of favorites to no greater than 30."
		'>>> Exit condition is the user having fewer than 30 scripts in their favorites menu.
	LOOP UNTIL ubound(double_check_array) <= 29

	'>>> Getting ready to write the user's selection to a text file and save it on a prescribed location on the network.
	'>>> Building the content of the text file.
	FOR i = 0 to number_of_scripts - 1
		IF scripts_edit_favs_array(i, 1) = checked THEN favorite_scripts = favorite_scripts & scripts_edit_favs_array(i, 2) & scripts_edit_favs_array(i, 3) & vbNewLine
	NEXT

	'>>> After the user selects their favorite scripts, we are going to write (or overwrite) the list of scripts
	'>>> stored at H:\my favorite scripts.txt.
	IF favorite_scripts <> "" THEN
		SET updated_fav_scripts_fso = CreateObject("Scripting.FileSystemObject")
		SET updated_fav_scripts_command = updated_fav_scripts_fso.CreateTextFile(favorites_text_file_location, 2)
		updated_fav_scripts_command.Write(favorite_scripts)
		updated_fav_scripts_command.Close
		script_end_procedure("Success!! Your Favorites Menu has been updated. Please click your favorites list button to re-load them.")
	ELSE
		'>>> OR...if the user has selected no scripts for their favorite, the file will be deleted to
		'>>> prevent the Favorites Menu from erroring out.
		'>>> Experience with worker_signature automation tells us that if the text file is blank, the favorites menu doth not work.
		oTxtFile.DeleteFile(favorites_text_file_location)
		script_end_procedure("You have updated your Favorites Menu, but you haven't selected any scripts. The next time you use the Favorites scripts, you will need to select your favorites.")
	END IF

end function

'>>> Determining the location of the user's favorites list.

'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

favorites_text_file_location = user_myDocs_folder & "\scripts-cs-favorites.txt"

'switching up the script_repository because the all scripts file is not in the Script Files folder
all_scripts_repo = script_repository & "/~complete-list-of-scripts.vbs"

'========================================================================================================
'========================================================================================================
'========================================================================================================
'========================================================================================================
'================================================================================== NOW THE ACTUAL SCRIPT
'========================================================================================================
'========================================================================================================
'========================================================================================================
'========================================================================================================

'>>> Our script arrays.
'>>> all_scripts_array will be built from the contents of the user's text file
'>>> new_scripts will be build automatically by looking at the description of each script in GitHub. If the description includes "NEW" then it is added to the array.
'>>> mandatory_array is pre-determined
all_scripts_array = ""
new_scripts = ""
'mandatory_array = "ACTIONS - NCP LOCATE" & vbNewLine & "ACTIONS - RECORD IW INFO" & vbNewLine & "ACTIONS - SEND F0104 DORD MEMO" & vbNewLine & "NOTES - ADJUSTMENTS" & vbNewLine & "NOTES - ARREARS MANAGEMENT REVIEW" & vbNewLine & "NOTES - CLIENT CONTACT"

'Does this differently if you're a run_locally user vs not
If run_locally <> true then
	'>>> Creating the object needed to connect to the interwebs.
	SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")
	'all_scripts_repo = script_repository & "~complete-list-of-scripts.vbs"
	get_all_scripts.open "GET", all_scripts_repo, FALSE
	get_all_scripts.send
	IF get_all_scripts.Status = 200 THEN
		Set filescriptobject = CreateObject("Scripting.FileSystemObject")
		Execute get_all_scripts.responseText
	ELSE
		'>>> Displaying the error message when the script fails to connect to a specific main menu.
		'>>> the replace & right bits are there to display the main menu in a way that is clear to the user.
		'>>> We are going to display the right length minus 99 because there are 99 characters between the start of the https and the last / before the main menu name.
		'>>> That length needs to be updated when we go state-wide.
		MsgBox("Something went wrong grabbing trying to locate All Scripts File. Please contact scripts administrator.")
		stopscript
	END IF
ELSE
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(all_scripts_repo)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
END IF

'>>> Building the array of new scripts
'>>> If the description of the script includes the word "NEW" then the script file name is added to the array.
num_of_new_scripts = 0
new_array = ""
FOR i = 0 TO Ubound(cs_scripts_array)
	IF DateDiff("D", cs_scripts_array(i).release_date, date) < 90 THEN new_array = new_array & UCASE(cs_scripts_array(i).category) & " - " & UCASE(replace(replace(cs_scripts_array(i).file_name, ".vbs", " "), "-", " ")) & vbNewLine
NEXT

'>>> Removing .vbs from the array for the prettification of the display to the users.
new_array = replace(new_array, ".vbs", "")

'>>> Custom function that builds the Favorites Main Menu dialog.
'>>> the array of the user's scripts
FUNCTION favorite_menu(user_scripts_array, mandatory_array, new_array, script_location, worker_signature)
	'>>> Splitting the array of all scripts.
	user_scripts_array = trim(user_scripts_array)
	user_scripts_array = split(user_scripts_array, vbNewLine)

	mandatory_array = trim(mandatory_array)
	mandatory_array = split(mandatory_array, vbNewLine)

	new_array = trim(new_array)
	new_array = split(new_array, vbNewLine)

	num_of_user_scripts = ubound(user_scripts_array)
	num_of_mandatory_scripts = ubound(mandatory_array)
	num_of_new_scripts = ubound(new_array)

	num_of_scripts = num_of_user_scripts + num_of_mandatory_scripts + num_of_new_scripts

	ReDim all_scripts_array(num_of_scripts, 5)
	'position 0 = script name
	'position 1 = script directory
	'position 2 = button
	'position 3 = category
	'position 4 = script name without category
	'position 5 = state-supported true/false

	scripts_pos = 0
	FOR EACH script_path IN user_scripts_array
		IF script_path <> "" THEN
			all_scripts_array(scripts_pos, 0) = script_path
			'>>> Creating the correct URL for the github call
			'>>> When we clean up this for state-wide deployment, we will need determine the appropriate network location for the agency custom scripts
			IF left(script_path, 5) = "notes" THEN
				all_scripts_array(scripts_pos, 1) = script_path
				all_scripts_array(scripts_pos, 3) = "NOTES"
				all_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 6)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_path, 7) = "actions" THEN
				all_scripts_array(scripts_pos, 1) = script_path
				all_scripts_array(scripts_pos, 3) = "ACTIONS"
				all_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 8)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_path, 4) = "bulk" THEN
				all_scripts_array(scripts_pos, 1) = script_path
				all_scripts_array(scripts_pos, 3) = "BULK"
				all_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 5)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_path, 11) = "calculators" THEN
				all_scripts_array(scripts_pos, 1) = script_path
				all_scripts_array(scripts_pos, 3) = "CALCULATORS"
				all_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 12)
				all_scripts_array(scripts_pos, 5) = true
            ELSEIF left(script_path, 9) = "utilities" THEN
    			all_scripts_array(scripts_pos, 1) = script_path
    			all_scripts_array(scripts_pos, 3) = "UTILITIES"
    			all_scripts_array(scripts_pos, 4) = right(script_path, len(script_path) - 10)
    			all_scripts_array(scripts_pos, 5) = true
			END IF
			scripts_pos = scripts_pos + 1
		END IF
	NEXT

	FOR EACH script_name IN mandatory_array
		IF script_name <> "" THEN
			all_scripts_array(scripts_pos, 0) = script_name
			'>>> Creating the correct URL for the github call
			'>>> When we clean up this for state-wide deployment, we will need determine the appropriate network location for the agency custom scripts
			IF left(script_name, 5) = "NOTES" THEN
				all_scripts_array(scripts_pos, 1) = "/NOTES/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "NOTES"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 7)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 7) = "ACTIONS" THEN
				all_scripts_array(scripts_pos, 1) = "/ACTIONS/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "ACTIONS"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 9)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 4) = "BULK" THEN
				all_scripts_array(scripts_pos, 1) = "/BULK/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "BULK"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 6)
				all_scripts_array(scripts_pos, 5) = true
			END IF
			scripts_pos = scripts_pos + 1
		END IF
	NEXT

	FOR EACH script_name IN new_array
		IF script_name <> "" THEN
			all_scripts_array(scripts_pos, 0) = script_name
			'>>> Creating the correct URL for the github call
			'>>> When we clean up this for state-wide deployment, we will need determine the appropriate network location for the agency custom scripts
			IF left(script_name, 5) = "NOTES" THEN
				all_scripts_array(scripts_pos, 1) = "/NOTES/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "NOTES"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 7)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 7) = "ACTIONS" THEN
				all_scripts_array(scripts_pos, 1) = "/ACTIONS/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "ACTIONS"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 9)
				all_scripts_array(scripts_pos, 5) = true
			ELSEIF left(script_name, 4) = "BULK" THEN
				all_scripts_array(scripts_pos, 1) = "/BULK/" & script_name & ".vbs"
				all_scripts_array(scripts_pos, 3) = "BULK"
				all_scripts_array(scripts_pos, 4) = right(script_name, len(script_name) - 6)
				all_scripts_array(scripts_pos, 5) = true
			END IF
			scripts_pos = scripts_pos + 1
		END IF
	NEXT

	'>>> Determining the height parameters to enable the group boxes.
	actions_count = 0
	bulk_count = 0
	calc_count = 0
	notes_count = 0
	FOR i = 0 TO (ubound(user_scripts_array) - 1)
		IF all_scripts_array(i, 3) = "ACTIONS" THEN
			actions_count = actions_count + 1
		ELSEIF all_scripts_array(i, 3) = "BULK" THEN
			bulk_count = bulk_count + 1
		ELSEIF all_scripts_array(i, 3) = "CALCULATORS" THEN
			calc_count = calc_count + 1
		ELSEIF all_scripts_array(i, 3) = "NOTES" THEN
			notes_count = notes_count + 1
        ELSEIF all_scripts_array(i, 3) = "UTILITIES" THEN
    		utilities_count = utilities_count + 1
		END IF
	NEXT

	'>>> Determining the height of the dialog.
	'>>> Each groupbox will require a minimum of 25 pixels. That is the height of the groupbox with 1 script PushButton
	'>>> The groupboxes need to grow 10 for each script pushbutton, so the dialog also needs to grow 10 for each script push button. However,
	'>>> 	the size of each groupbox will always be 15 plus (10 times the number of that kind of script)...
	dlg_height = 0
	IF actions_count <> 0 THEN dlg_height = 15 + (10 * actions_count)
	IF bulk_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (10 * bulk_count))
	IF calc_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (10 * bulk_count))
	IF notes_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (10 * notes_count))
    IF utilities_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (10 * utilities_count))
	dlg_height = dlg_height + 5
	'>>> The dialog needs to be at least 185 pixels tall. If it is not...because the user has not selected a sufficient number of scripts...then
	'>>> the script needs to grow to 185.

	'>>> Adjusting the height if the user has fewer scripts than what is "recommended" plus the new scripts
	alt_dlg_height = 60 + (10 * (Ubound(mandatory_array) + 1)) + (10 * (Ubound(new_array) + 1))
	IF alt_dlg_height > dlg_height THEN dlg_height = alt_dlg_height

	'>>> Determining the start row for the push buttons
	'>>> The position of one groupbox will be determined from the existence of other groupboxes earlier in the alphabet.
	'>>> The actions start row is 10, and the end row will be 10 plus 15 (for the default height of the groupbox) plus 10 for each ACTIONS script
	IF actions_count <> 0 THEN
		actions_start_row = 10
		actions_end_row = 10 + (15 + (10 * actions_count))
	ELSE
		'>>> ...or they will both be 0 when there are not ACTIONS scripts in the user's favorites.
		actions_start_row = 0
		actions_end_row = 0
	END IF
	'>>> The BULK groupbox start row will be determined by the end of the ACTIONS row...and so on.
	IF bulk_count <> 0 THEN
		bulk_start_row = 10 + actions_end_row
		bulk_end_row = bulk_start_row + (15 + (10 * bulk_count))
	ELSE
		bulk_start_row = actions_start_row
		bulk_end_row = actions_end_row
	END IF
	IF calc_count <> 0 THEN
		calc_start_row = 10 + bulk_end_row
		calc_end_row = calc_start_row + (15 + (10 * calc_count))
	ELSE
		calc_start_row = bulk_start_row
		calc_end_row = bulk_end_row
	END IF
	IF notes_count <> 0 THEN
		notes_start_row = 10 + calc_end_row
		notes_end_row = notes_start_row + (15 + (10 * notes_count))
	ELSE
		notes_start_row = calc_start_row
		notes_end_row = calc_end_row
	END IF
    IF utilities_count <> 0 THEN
		utilities_start_row = 10 + notes_end_row
		utilities_end_row = utilities_start_row + (15 + (10 * utilities_count))
	ELSE
		utilities_start_row = notes_start_row
		utilities_end_row = notes_end_row
	END IF

	'>>> A nice decoration for the user. If they have used Update Worker Signature, then their signature is built into the dialog display.
	IF worker_signature <> "" THEN
		dlg_name = worker_signature & "'s Favorite Scripts"
	ELSE
		dlg_name = "My Favorite Scripts"
	END IF

	'>>> The dialog
	BeginDialog favorites_dialog, 0, 0, 411, dlg_height, dlg_name & " "
  	  ButtonGroup ButtonPressed
		'>>> User's favorites
		'>>> Here, we are using the value for the script type start_row to determine the vertical position of each pushbutton.
		'>>> As we add a pushbutton, we need to increase the value for the start_row by 10 for that kind of script.
		FOR i = 0 TO (ubound(user_scripts_array) - 1)
			IF all_scripts_array(i, 3) = "ACTIONS" THEN
				PushButton 20, actions_start_row + 10, 170, 10, all_scripts_array(i, 4), all_scripts_array(i, 2)
				actions_start_row = actions_start_row + 10
			ELSEIF all_scripts_array(i, 3) = "BULK" THEN
				PushButton 20, bulk_start_row + 10, 170, 10, all_scripts_array(i, 4), all_scripts_array(i, 2)
				bulk_start_row = bulk_start_row + 10
			ELSEIF all_scripts_array(i, 3) = "CALCULATORS" THEN
				PushButton 20, calc_start_row + 10, 170, 10, all_scripts_array(i, 4), all_scripts_array(i, 2)
				calc_start_row = calc_start_row + 10
			ELSEIF all_scripts_array(i, 3) = "NOTES" THEN
				PushButton 20, notes_start_row + 10, 170, 10, all_scripts_array(i, 4), all_scripts_array(i, 2)
				notes_start_row = notes_start_row + 10
            ELSEIF all_scripts_array(i, 3) = "UTILITIES" THEN
				PushButton 20, utilities_start_row + 10, 170, 10, all_scripts_array(i, 4), all_scripts_array(i, 2)
				utilities_start_row = utilities_start_row + 10
			END IF
		NEXT

		'>>> Placing Mandatory Scripts
		FOR i = ubound(user_scripts_array) to (ubound(user_scripts_array) + (ubound(mandatory_array) - 1))
			right_hand_row = (20 + (10 * (i - num_of_user_scripts)))
			PushButton 220, right_hand_row, 180, 10, all_scripts_array(i, 0), all_scripts_array(i, 2)
		NEXT

		right_hand_row = right_hand_row + 30
		'>>> Placing new scripts
		FOR i = (ubound(user_scripts_array) + ubound(mandatory_array)) to (ubound(user_scripts_array) + ubound(mandatory_array) + (ubound(new_array) - 1))
			PushButton 220, right_hand_row, 180, 10, all_scripts_array(i, 0), all_scripts_array(i, 2)
			right_hand_row = right_hand_row + 10
		NEXT

		'>>> Placing groupboxes.
		'>>> All of the objects need to be placed at the end of the dialog. If they are not, it will throw off the positioning of the PushButtons
		'>>> which will, in turn, throw off the calculations for which script should be run.
		'>>> The height and position of each GroupBox is determed dynamically from the number of scripts in the groups previous.
		'>>> Mandatory and New are always going to be in the there, and located on the right hand side of the DLG.
        GroupBox 210, 10, 195, 5 + (10 * (Ubound(mandatory_array) + 1)), "Recommended Scripts"
		GroupBox 210, 20 + (10 * (Ubound(mandatory_array) + 1)), 195, 5 + (10 * (UBound(new_array) + 1)), "NEW SCRIPTS!!!"
		IF actions_count <> 0 THEN GroupBox 5, 10, 195, (15 + (10 * actions_count)), "ACTIONS"
		IF bulk_count <> 0 THEN GroupBox 5, actions_end_row + 10, 195, (15 + (10 * bulk_count)), "BULK"
		IF calc_count <> 0 THEN GroupBox 5, bulk_end_row + 10, 195, (15 + (10 * calc_count)), "CALCULATORS"
		IF notes_count <> 0 THEN GroupBox 5, calc_end_row + 10, 195, (15 + (10 * notes_count)), "NOTES"
        IF utilities_count <> 0 THEN GroupBox 5, notes_end_row + 10, 195, (15 + (10 * utilities_count)), "UTILITIES"
		PushButton 210, dlg_height - 25, 65, 15, "Update Favorites", update_favorites_button
		PushButton 285, dlg_height - 25, 60, 15, "Update Hotkeys", update_hotkeys_button
		CancelButton 355, dlg_height - 25, 50, 15
	EndDialog

	'>>> Loading the favorites dialog
	DIALOG favorites_dialog
	'>>> Cancelling the script if ButtonPressed = 0
	IF ButtonPressed = 0 THEN stopscript
	'>>> Giving user has the option of updating their favorites menu.
	'>>> We should try to incorporate the chainloading function of the new script_end_procedure to bring the user back to their favorites.
	IF buttonpressed = update_favorites_button THEN
		call edit_favorites
		StopScript
	ElseIf buttonpressed = update_hotkeys_button then
		call edit_hotkeys
		StopScript
	End if
	'>>> This tells the script which PushButton has been selected.
	'>>> We need to do ButtonPressed - 1 because of the way that the system assigns a value to ButtonPressed.
	'>>> When then favorites menu is launched from the Powerpad, the formula is ButtonPressed - 1. But if the menu is hidden behind another menu, then this formula is ButtonPressed - 1 - the number of other buttons ahead of the favorites menu button in that dialog tab order.
	script_location = all_scripts_array(ButtonPressed - 1, 1)  '!!!! THIS WILL NEED TO BE buttonpressed - (the number of objects created before the PushButtons...which is the dialog itself. don't move the order of the pushbuttons!!
	script_location = lcase(script_location)
	script_location = replace(script_location, " ", "-")
	script_location = replace(script_location, "bulk:-", "")
	script_location = replace(script_location, "actions:-", "")
	script_location = replace(script_location, "notes:-", "")
	script_location = replace(script_location, "notes---", "")
	script_location = replace(script_location, "actions---", "")
	script_location = replace(script_location, "bulk---", "")
	script_location = replace(script_location, "calc:-", "")
	script_location = replace(script_location, "calc---", "")
    script_location = replace(script_location, "utilities:-", "")
	script_location = replace(script_location, "utilities---", "")
END FUNCTION
'======================================

'The script starts HERE!!!-------------------------------------------------------------------------------------------------------------------------------------

'>>> The gobbins of the script that the user sees and makes do.
'>>> Declaring the text file storing the user's favorite scripts list.
Dim oTxtFile
With (CreateObject("Scripting.FileSystemObject"))
	'>>> If the file exists, we will grab the list of the user's favorite scripts and run the favorites menu.
	If .FileExists(favorites_text_file_location) Then
		Set fav_scripts = CreateObject("Scripting.FileSystemObject")
		Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
		fav_scripts_array = fav_scripts_command.ReadAll
		IF fav_scripts_array <> "" THEN user_scripts_array = fav_scripts_array
		fav_scripts_command.Close
	ELSE
		'>>> ...otherwise, if the file does not exist, the script will require the user to select their favorite scripts.
		call edit_favorites
	END IF
END WITH

'>>> Calling the function that builds the favorites menu.
CALL favorite_menu(user_scripts_array, mandatory_array, new_array, script_location, worker_signature)

'>>> Running the script that is selected.
'>>> The first determination is whether the script is located on the agency's network.
'>>> Running the script if it is agency-custom script

'>>> Running the script
script_URL = script_repository & script_location
If run_locally = true then
    Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
    Set fso_command = run_another_script_fso.OpenTextFile(script_URL)
    text_from_the_other_script = fso_command.ReadAll
    fso_command.Close
    Execute text_from_the_other_script
Else
    CALL run_from_GitHub(script_URL)
End if




'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< AGAIN, VERY TEMPORARY
'END IF
